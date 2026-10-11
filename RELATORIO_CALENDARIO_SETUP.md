# APS PURAN - Calendario, dias de producao e Setup

**Nada foi compilado no VBA nem testado no Excel.** Verificacoes feitas: leitura/analise estatica do codigo, varredura de variaveis nao declaradas, equilibrio de blocos, ASCII/CRLF, e uma simulacao externa em Python (so aritmetica de calendario).

## Arquivos
| Arquivo | Situacao |
|---|---|
| frmAPS_Config_codigo.txt | NOVO/reescrito: formulario do calendario (codigo completo; controles criados em codigo) |
| modAPS_Calendario.bas | MODIFICADO (leitura dos horarios por dia; excecao sem horario usa o horario do dia da semana; CR solto removido) |
| modAPS_Base.bas | MODIFICADO (APS_MontarAuxDaColuna Public; coluna Setup editada = bit 64) |
| modAPS_Motor.bas | MODIFICADO (APS_ReencaixarLinha: edicao do Setup na coluna mantem o inicio da preparacao) |
| frmAPS_Operacao_codigo.txt | MODIFICADO (horario = inicio da PREPARACAO; previa com Setup) |
| modAPS_Cards.bas | MODIFICADO antes (atraso, botao Calendario com tratamento de erro) |
| modAPS_SAP.bas, EstaPastaDeTrabalho | NAO modificados |

## Causas / mudancas
1. **Botao Calendario**: pela leitura, Botao -> APS_AbrirCalendario -> frmAPS_Config.Show esta correto. A causa real NAO foi confirmada; suspeita: projeto sem compilar (APS_PararTick duplicado no SEU projeto). Adicionei tratamento de erro que mostra a mensagem real. Se persistir, envie o frmAPS_Config atual.
2. **Dias da semana independentes + datas especificas**: novo formulario com 7 dias independentes (marcar/horario) e calendario mensal (clique no dia: Padrao / Liberar / Bloquear). Excecoes valem so para a data completa (dd/mm/aaaa), tem prioridade sobre o padrao so nesse dia, ficam em CAL_EXC/CAL_SEM (persistem). O defeito "alteracao afeta outros dias" NAO foi reproduzido no codigo do motor nem do formulario; nao posso afirmar que foi a causa.
3. **Setup**: o campo de horario do formulario era o INICIO DA PRODUCAO (a preparacao era calculada para tras). Agora, com preparacao marcada, o campo e o INICIO DA PREPARACAO e a producao = inicio + Setup (+Limpeza) pelo calendario. Ex.: 10:00 + 1:30 -> 11:30. Aumentar o setup atrasa a producao; reduzir antecipa e a cascata fecha o vao (so tempo liberado por evento; folgas intencionais ficam). Na planilha, editar a coluna Setup mantem o inicio da preparacao. Coluna Inicio da planilha continua sendo o inicio da PRODUCAO. Duracao e quantidade originais preservadas; sem card/linha duplicados (o Setup continua virtual, vinculado a producao).
4. Se o usuario escolher "Nenhuma" ao editar, a producao passa a comecar no horario mostrado (inicio da preparacao).

## Testes
- Simulacao externa (simulacao_setup.py): 10:00+1:00=11:00; +1:30=11:30; +2:00=12:00; +0:30=10:30; virada de turno ok. Executada.
- Compilacao VBA: NAO executada. Excel real: NAO executado.

## Teste sugerido no Excel
Setup 1h -> 2h (producao anda +1h, nao antes); 2h -> 1h (antecipa); preparacao 10:00 + 1:30 (producao 11:30, card Setup + producao); calendario: desmarcar so a quinta, liberar um sabado especifico, fechar/reabrir; mes seguinte sem efeito.

## Limitacoes
Formulario do calendario nao compilado; datas fora do mes limite seguem APS_LimiteMes; Limpeza entra no tempo de preparacao junto com o Setup.

## Instalacao
Backup (salvar copia .xlsm). Remover e importar modAPS_Base, modAPS_Calendario, modAPS_Motor, modAPS_Cards (Arquivo > Importar). Substituir TODO o codigo de frmAPS_Operacao e frmAPS_Config pelo texto dos .txt. NAO mexer em modAPS_SAP. Depurar > Compilar VBAProject; reportar mensagem e linha. Restaurar: reabrir a copia de backup.
