from __future__ import annotations

import io
import csv
import re
import sqlite3
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import BinaryIO

from flask import Flask, flash, redirect, render_template, request, url_for


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_REPORT = BASE_DIR / "01-Relatorio.xls"
CURRENT_REPORT = BASE_DIR / "current_report.xls"
AUDIT_DB = BASE_DIR / "ponto_audit.sqlite3"
CURRENT_CSV = BASE_DIR / "ponto_dados.csv"
SS_NS = "urn:schemas-microsoft-com:office:spreadsheet"
NS = {"ss": SS_NS}
TIME_PATTERN = re.compile(r"\b([01]?\d|2[0-3]):([0-5]\d)\b")

app = Flask(__name__)
app.secret_key = "gui-ponto-local"


@dataclass
class DayResult:
    day: date
    marks: list[str]
    worked_minutes: int
    expected_minutes: int
    balance_minutes: int
    original_marks: list[str]
    deviations: list[int | None]
    deviation_minutes: float
    heat_level: int = 0
    audit_reason: str | None = None
    day_type: str = "útil"
    day_label: str = ""
    morning_percent: int = 0
    warning: str | None = None


@dataclass
class EmployeeResult:
    employee_id: str
    name: str
    department: str
    days: list[DayResult]

    @property
    def worked_minutes(self) -> int:
        return sum(day.worked_minutes for day in self.days)

    @property
    def expected_minutes(self) -> int:
        return sum(day.expected_minutes for day in self.days)

    @property
    def balance_minutes(self) -> int:
        return self.worked_minutes - self.expected_minutes

    @property
    def attendance_ratio(self) -> float:
        return self.worked_minutes / self.expected_minutes if self.expected_minutes else 1

    @property
    def eligible(self) -> bool:
        return self.worked_minutes > self.expected_minutes * 0.5


def _cell_values(row: ET.Element) -> list[str]:
    values: list[str] = []
    for cell in row.findall("ss:Cell", NS):
        index = cell.attrib.get(f"{{{SS_NS}}}Index")
        if index:
            values.extend([""] * max(0, int(index) - len(values) - 1))
        data = cell.find("ss:Data", NS)
        values.append((data.text or "").strip() if data is not None else "")
    return values


def _rows(worksheet: ET.Element) -> list[list[str]]:
    table = worksheet.find("ss:Table", NS)
    if table is None:
        return []
    return [_cell_values(row) for row in table.findall("ss:Row", NS)]


def _parse_period(rows: list[list[str]]) -> tuple[int, int]:
    for row in rows:
        value = " ".join(row)
        match = re.search(r"Data:\s*(\d{4})\.(\d{2})\.\d{2}~", value)
        if match:
            return int(match.group(1)), int(match.group(2))
    raise ValueError("Não foi possível encontrar o período do relatório.")


def _parse_time(value: str) -> time | None:
    match = TIME_PATTERN.search(value)
    if not match:
        return None
    return time(int(match.group(1)), int(match.group(2)))


def _format_minutes(minutes: int, signed: bool = False) -> str:
    sign = ""
    if signed and minutes < 0:
        sign, minutes = "-", -minutes
    hours, remainder = divmod(minutes, 60)
    return f"{sign}{hours:02d}:{remainder:02d}"


def _easter_sunday(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def _brazilian_optional_holidays(year: int) -> dict[date, str]:
    easter = _easter_sunday(year)
    return {
        date(year, 1, 1): "Confraternização Universal",
        easter - timedelta(days=48): "Carnaval (ponto facultativo)",
        easter - timedelta(days=47): "Carnaval (ponto facultativo)",
        easter - timedelta(days=2): "Sexta-feira Santa",
        easter - timedelta(days=1): "Sábado de Aleluia",
        easter: "Domingo de Páscoa",
        easter + timedelta(days=60): "Corpus Christi (ponto facultativo)",
        date(year, 4, 21): "Tiradentes",
        date(year, 5, 1): "Dia do Trabalho",
        date(year, 9, 7): "Independência do Brasil",
        date(year, 10, 12): "Nossa Senhora Aparecida",
        date(year, 11, 2): "Finados",
        date(year, 11, 15): "Proclamação da República",
        date(year, 11, 20): "Dia da Consciência Negra",
        date(year, 12, 25): "Natal",
    }


def _day_classification(current_day: date, holidays: dict[date, str]) -> tuple[str, str, int]:
    if current_day in holidays:
        return "feriado", holidays[current_day], 0
    if current_day.weekday() == 5:
        return "sábado", "Sábado (facultativo)", 0
    if current_day.weekday() == 6:
        return "domingo", "Domingo (facultativo)", 0
    return "útil", "Dia útil (obrigatório)", 1


def _database() -> sqlite3.Connection:
    connection = sqlite3.connect(AUDIT_DB)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_edits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT NOT NULL,
            day TEXT NOT NULL,
            original_marks TEXT NOT NULL,
            edited_marks TEXT NOT NULL,
            reason TEXT NOT NULL,
            changed_at TEXT NOT NULL
        )
        """
    )
    return connection


def _audit_marks(
    employee_id: str, current_day: date, marks: list[str]
) -> tuple[list[str], list[str] | None, str | None]:
    with _database() as connection:
        row = connection.execute(
            """
            SELECT original_marks, edited_marks, reason
            FROM audit_edits
            WHERE employee_id = ? AND day = ?
            ORDER BY id DESC LIMIT 1
            """,
            (employee_id, current_day.isoformat()),
        ).fetchone()
    if not row:
        return marks, None, None
    edited = [value for value in row[1].split(",") if value]
    original = [value for value in row[0].split(",") if value]
    return edited, original, row[2]


def _save_audit(
    employee_id: str, current_day: date, original_marks: list[str], edited_marks: list[str], reason: str
) -> None:
    with _database() as connection:
        connection.execute(
            """
            INSERT INTO audit_edits
              (employee_id, day, original_marks, edited_marks, reason, changed_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                employee_id,
                current_day.isoformat(),
                ",".join(original_marks),
                ",".join(edited_marks),
                reason,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )


def _valid_marks(raw_marks: list[str]) -> list[str]:
    normalized = {
        f"{int(value.split(':')[0]):02d}:{int(value.split(':')[1]):02d}"
        for value in raw_marks
    }
    unique = sorted(normalized, key=lambda value: (int(value[:2]), int(value[3:])))
    return unique[:4]


def _day_result(
    year: int,
    month: int,
    day_number: int,
    raw_marks: str,
    regular_minutes: int,
    employee_id: str,
    holidays: dict[date, str],
) -> DayResult:
    current_day = date(year, month, day_number)
    day_type, day_label, is_required = _day_classification(current_day, holidays)
    original_marks = _valid_marks([match.group(0) for match in TIME_PATTERN.finditer(raw_marks)])
    marks, audited_original, audit_reason = _audit_marks(
        employee_id, current_day, original_marks
    )
    if audited_original is not None:
        original_marks = audited_original
    parsed_marks = [_parse_time(mark) for mark in marks]
    valid_marks = [mark for mark in parsed_marks if mark is not None]
    worked = 0
    for start, end in zip(valid_marks[::2], valid_marks[1::2]):
        start_minutes = start.hour * 60 + start.minute
        end_minutes = end.hour * 60 + end.minute
        if end_minutes >= start_minutes:
            worked += end_minutes - start_minutes

    expected = regular_minutes if is_required else 0
    warning = None
    if len(valid_marks) % 2:
        warning = "Quantidade ímpar de marcações; a última foi desconsiderada."
    nominal = [8 * 60, 12 * 60, 13 * 60, 18 * 60]
    deviations: list[int | None] = []
    for index in range(4):
        if index < len(valid_marks):
            value = valid_marks[index].hour * 60 + valid_marks[index].minute
            deviations.append(value - nominal[index])
        else:
            deviations.append(None)
    absolute_deviations = [abs(value) for value in deviations if value is not None]
    return DayResult(
        day=current_day,
        marks=marks,
        original_marks=original_marks,
        deviations=deviations,
        deviation_minutes=sum(absolute_deviations) / len(absolute_deviations)
        if absolute_deviations
        else 0,
        audit_reason=audit_reason,
        day_type=day_type,
        day_label=day_label,
        morning_percent=_morning_percent(marks),
        worked_minutes=worked,
        expected_minutes=expected,
        balance_minutes=worked - expected,
        warning=warning,
    )


def _morning_percent(marks: list[str]) -> int:
    if not marks:
        return 0
    entry = _parse_time(marks[0])
    if entry is None:
        return 0
    actual = entry.hour * 60 + entry.minute
    nominal = 8 * 60
    return max(0, min(100, round(100 - max(0, actual - nominal) / 480 * 100)))


def _save_csv(employees: list[EmployeeResult]) -> None:
    with CURRENT_CSV.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(
            output,
            fieldnames=["Data", "ID", "Nome", "Departamento", "Entrada", "Almoço", "Retorno", "Saida"],
        )
        writer.writeheader()
        for employee in employees:
            for day in employee.days:
                marks = day.marks + [""] * (4 - len(day.marks))
                writer.writerow(
                    {
                        "Data": day.day.isoformat(),
                        "ID": employee.employee_id,
                        "Nome": employee.name,
                        "Departamento": employee.department,
                        "Entrada": marks[0],
                        "Almoço": marks[1],
                        "Retorno": marks[2],
                        "Saida": marks[3],
                    }
                )


def _load_csv() -> list[EmployeeResult]:
    grouped: dict[tuple[str, str, str], list[dict[str, str]]] = {}
    with CURRENT_CSV.open("r", encoding="utf-8-sig", newline="") as source:
        for row in csv.DictReader(source):
            key = (row.get("ID", ""), row.get("Nome", ""), row.get("Departamento", ""))
            grouped.setdefault(key, []).append(row)
    if not grouped:
        raise ValueError("O CSV de ponto está vazio.")
    first_row = next(iter(grouped.values()))[0]
    first_date = date.fromisoformat(first_row["Data"])
    holidays = _brazilian_optional_holidays(first_date.year)
    employees = []
    for (employee_id, name, department), rows in grouped.items():
        days = []
        for row in rows:
            current_day = date.fromisoformat(row["Data"])
            raw_marks = "\n".join(row.get(column, "") for column in ("Entrada", "Almoço", "Retorno", "Saida"))
            days.append(
                _day_result(
                    current_day.year,
                    current_day.month,
                    current_day.day,
                    raw_marks,
                    8 * 60,
                    employee_id,
                    holidays,
                )
            )
        employees.append(EmployeeResult(employee_id, name, department, days))
    return _finish_calculations(employees)


def _finish_calculations(employees: list[EmployeeResult]) -> list[EmployeeResult]:
    for employee in employees:
        average = sum(day.deviation_minutes for day in employee.days if day.deviation_minutes)
        days_with_deviation = sum(bool(day.deviation_minutes) for day in employee.days)
        average = average / days_with_deviation if days_with_deviation else 0
        for day in employee.days:
            if not day.deviation_minutes or not average:
                day.heat_level = 0
            elif day.deviation_minutes <= average * 0.5:
                day.heat_level = 0
            elif day.deviation_minutes <= average:
                day.heat_level = 1
            elif day.deviation_minutes <= average * 1.5:
                day.heat_level = 2
            else:
                day.heat_level = 3
    return employees


def _regular_minutes(root: ET.Element) -> int:
    schedule_text = next(
        (
            data.text or ""
            for data in root.findall(".//ss:Data", NS)
            if "Turno" in (data.text or "")
        ),
        "",
    )
    times = [
        _parse_time(match.group(0))
        for match in TIME_PATTERN.finditer(schedule_text)
    ]
    valid_times = [value for value in times if value is not None]
    total = 0
    for start, end in zip(valid_times[::2], valid_times[1::2]):
        start_minutes = start.hour * 60 + start.minute
        end_minutes = end.hour * 60 + end.minute
        if end_minutes >= start_minutes:
            total += end_minutes - start_minutes
    return total or 8 * 60


def parse_report(source: BinaryIO | bytes) -> list[EmployeeResult]:
    if isinstance(source, bytes):
        source = io.BytesIO(source)
    try:
        root = ET.parse(source).getroot()
    except ET.ParseError as error:
        raise ValueError("O arquivo não é um relatório XML válido do relógio de ponto.") from error

    detail_sheet = next(
        (
            worksheet
            for worksheet in root.findall(".//ss:Worksheet", NS)
            if "detalhe" in worksheet.attrib.get(f"{{{SS_NS}}}Name", "").lower()
        ),
        None,
    )
    if detail_sheet is None:
        raise ValueError("A aba 'Detalhe' não foi encontrada no relatório.")

    rows = _rows(detail_sheet)
    year, month = _parse_period(rows)
    regular_minutes = _regular_minutes(root)
    holidays = _brazilian_optional_holidays(year)
    employees: list[EmployeeResult] = []
    for index, row in enumerate(rows):
        if not row or "ID:" not in row:
            continue
        try:
            id_index = row.index("ID:")
            name_index = row.index("Nome:")
            department_index = row.index("Dept:")
            employee_id = row[id_index + 1]
            name = row[name_index + 1]
            department = row[department_index + 1]
        except (ValueError, IndexError):
            continue
        marks_row = rows[index + 1] if index + 1 < len(rows) else []
        days: list[DayResult] = []
        for day_number in range(1, 32):
            raw_marks = marks_row[day_number - 1] if day_number <= len(marks_row) else ""
            try:
                days.append(
                    _day_result(
                        year,
                        month,
                        day_number,
                        raw_marks,
                        regular_minutes,
                        employee_id,
                        holidays,
                    )
                )
            except ValueError:
                continue
        days = [day for day in days if day.day.month == month]
        employees.append(EmployeeResult(employee_id, name, department, days))

    if not employees:
        raise ValueError("Nenhum funcionário foi encontrado na aba 'Detalhe'.")
    for employee in employees:
        average = sum(day.deviation_minutes for day in employee.days if day.deviation_minutes)
        days_with_deviation = sum(bool(day.deviation_minutes) for day in employee.days)
        average = average / days_with_deviation if days_with_deviation else 0
        for day in employee.days:
            if not day.deviation_minutes or not average:
                day.heat_level = 0
            elif day.deviation_minutes <= average * 0.5:
                day.heat_level = 0
            elif day.deviation_minutes <= average:
                day.heat_level = 1
            elif day.deviation_minutes <= average * 1.5:
                day.heat_level = 2
            else:
                day.heat_level = 3
    return employees


def _read_default_report() -> bytes:
    if not DEFAULT_REPORT.exists():
        raise ValueError("O arquivo 01-Relatorio.xls não foi encontrado na pasta da aplicação.")
    return DEFAULT_REPORT.read_bytes()


def _read_current_report() -> tuple[bytes, str]:
    if CURRENT_REPORT.exists():
        return CURRENT_REPORT.read_bytes(), CURRENT_REPORT.name
    return _read_default_report(), DEFAULT_REPORT.name


def _update_csv_row(employee_id: str, day: date, marks: list[str]) -> list[str]:
    with CURRENT_CSV.open("r", encoding="utf-8-sig", newline="") as source:
        rows = list(csv.DictReader(source))
    with CURRENT_CSV.open("r", encoding="utf-8-sig", newline="") as source:
        fieldnames = csv.DictReader(source).fieldnames
    if not fieldnames:
        raise ValueError("O CSV de ponto não possui cabeçalho.")
    columns = ("Entrada", "Almoço", "Retorno", "Saida")
    for row in rows:
        if row.get("ID") == employee_id and row.get("Data") == day.isoformat():
            for index, column in enumerate(columns):
                row[column] = marks[index] if index < len(marks) else ""
            break
    else:
        raise ValueError("Funcionário ou data não encontrado no CSV.")
    with CURRENT_CSV.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    return marks


@app.template_filter("minutes")
def minutes_filter(value: int) -> str:
    return _format_minutes(value, signed=True)


@app.template_filter("date_br")
def date_br(value: date) -> str:
    return value.strftime("%d/%m/%Y")


@app.template_filter("deviation")
def deviation_filter(value: int | None) -> str:
    return "—" if value is None else _format_minutes(value, signed=True)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded = request.files.get("report")
        if uploaded is None or not uploaded.filename:
            flash("Selecione um arquivo de relatório.", "error")
            return redirect(url_for("index"))
        try:
            CURRENT_REPORT.write_bytes(uploaded.read())
            with _database() as connection:
                connection.execute("DELETE FROM audit_edits")
            employees = parse_report(CURRENT_REPORT.read_bytes())
            _save_csv(employees)
            filename = uploaded.filename
        except ValueError as error:
            flash(str(error), "error")
            return redirect(url_for("index"))
    else:
        try:
            if CURRENT_CSV.exists():
                employees = _load_csv()
                filename = CURRENT_CSV.name
            else:
                report, filename = _read_current_report()
                employees = parse_report(report)
                _save_csv(employees)
        except ValueError as error:
            flash(str(error), "error")
            employees, filename = [], ""
    selected_id = request.args.get("employee", "all")
    eligible_employees = [employee for employee in employees if employee.eligible]
    visible_employees = (
        eligible_employees
        if selected_id == "all"
        else [employee for employee in eligible_employees if employee.employee_id == selected_id]
    )
    excluded_count = len(employees) - len(visible_employees)
    return render_template(
        "index.html",
        employees=visible_employees,
        filename=filename,
        excluded_count=excluded_count,
        all_employees=employees,
        selected_id=selected_id,
    )


@app.post("/edit")
def edit_day():
    employee_id = request.form.get("employee_id", "").strip()
    day_text = request.form.get("day", "").strip()
    marks = [
        request.form.get("entrada", "").strip(),
        request.form.get("almoco", "").strip(),
        request.form.get("retorno", "").strip(),
        request.form.get("saida", "").strip(),
    ]
    reason = request.form.get("reason", "").strip()
    try:
        current_day = date.fromisoformat(day_text)
        edited_marks = [
            (
                f"{int(mark.split(':')[0]):02d}:{int(mark.split(':')[1]):02d}"
                if TIME_PATTERN.fullmatch(mark)
                else ""
            )
            for mark in marks
        ]
        report, _ = _read_current_report()
        employees = parse_report(report)
        employee = next(item for item in employees if item.employee_id == employee_id)
        day = next(item for item in employee.days if item.day == current_day)
        if not reason:
            raise ValueError("Informe o motivo da alteração.")
        _save_audit(employee_id, current_day, day.marks, edited_marks, reason)
        _update_csv_row(employee_id, current_day, edited_marks)
        flash("Alteração salva na auditoria.", "success")
    except (ValueError, StopIteration) as error:
        flash(str(error) or "Não foi possível salvar a alteração.", "error")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
