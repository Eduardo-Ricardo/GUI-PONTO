# GUI-PONTO

Aplicação web para ler o arquivo bruto `01-Relatorio.xls` (formato XML exportado
pelo relógio de ponto) e mostrar horas trabalhadas, horas previstas e saldo por
funcionário.

## Executar

1. Instale as dependências: `py -m pip install -r requirements.txt`
2. Inicie o servidor: `py app.py`
3. Abra `http://127.0.0.1:5000`

O programa carrega automaticamente o `01-Relatorio.xls` existente. Para
atualizar os dados, exporte um novo relatório no mesmo formato e envie-o pelo
campo **Processar relatório**. O arquivo não precisa ser renomeado.

A jornada regular usada no cálculo é de 08:00 por dia útil; sábados e domingos
têm jornada prevista de 00:00. As marcações são agrupadas em pares
entrada/saída. Marcações sem par são exibidas com um aviso.

Antes do cálculo, o relatório é convertido para `ponto_dados.csv`, com uma
linha por funcionário e data. Ele contém `Nome`, `Entrada`, `Almoço`, `Retorno`
e `Saida`, além de `Data`, `ID` e `Departamento`. A tela passa a ler esse CSV;
alterações feitas na grade atualizam o próprio arquivo.

Cada dia do calendário é dividido em duas metades. A parte superior mostra a
barra da manhã: 100% corresponde à entrada nominal às 08:00 e a barra diminui
quando a entrada fica mais tarde. A parte inferior mostra as quatro batidas e
o desvio do dia. O seletor permite ver todos os funcionários elegíveis ou
apenas um.

O calendário começa no domingo e funciona como uma grade de células. Cada dia
possui um editor direto para as quatro batidas e o motivo da alteração. Ao
salvar, a linha correspondente é atualizada no CSV e a tela é recalculada.

O calendário respeita a regra brasileira adotada no painel: segunda a sexta
são dias úteis obrigatórios; sábado e domingo são facultativos; feriados
nacionais e pontos facultativos ficam marcados no calendário e não geram horas
previstas. A lista inclui feriados fixos e datas móveis calculadas para o ano
do relatório.

## Visualização e auditoria

O painel principal remove funcionários que trabalharam **50% ou menos** da
carga útil prevista no mês. Cada funcionário restante possui um calendário de
mapa de calor baseado nas quatro batidas nominais `08:00`, `12:00`, `13:00` e
`18:00`. Verde representa os dias mais próximos da escala e vermelho os dias
com maior desvio, comparados com a média mensal daquele funcionário.

Em **Ver detalhamento diário > Editar**, informe as batidas corrigidas e o
motivo. A alteração é gravada em `ponto_audit.sqlite3`, o valor original é
preservado e fica visível junto ao motivo da auditoria. Ao enviar um novo
relatório, a auditoria anterior é limpa para não misturar pessoas e datas de
relatórios diferentes.