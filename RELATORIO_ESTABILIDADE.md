# APS PURAN - Correcoes de estabilidade (codigo apenas)

Base: correcao_motor. Alterados: modAPS_Motor, modAPS_Cards, modAPS_Base. Demais modulos identicos.
STATUS: NAO compilado e NAO testado no Excel. Importar numa COPIA e rodar Depurar > Compilar VBAProject.

## Itens do relatorio mestre x codigo real
| Item | Situacao | Acao |
|---|---|---|
| Tick duplicado | CONFIRMADO (APS_AgendarTick nao cancelava o anterior) | Cancela mProx antes de reagendar; Tick sempre reagenda |
| Tick engolia erros (On Error Resume Next global) | CONFIRMADO | On Error GoTo, erro vai a barra de status, reagenda sempre |
| APS_AtualizarPlanejamento com pop-up | CONFIRMADO | Chama APS_Atualizar True (silencioso; erros ainda aparecem) |
| Edicoes descartadas em colagem grande (500 linhas / 3000 celulas) | CONFIRMADO | Sem limite: bits por numero de linha |
| AoAbrir com Resume Next total | CONFIRMADO | Etapas isoladas e falhas reportadas em um unico aviso |
| ScreenUpdating forcado para True | CONFIRMADO (menor) | Restaura o estado original |
| Status manual "Atrasada" revertido pelo automatico | NAO CONFIRMADO: StatusAutomatico ja exclui Atrasada/Cancelada | nada |
| APS_LimparSetupsOrfaos / APS_Ocupado | Incremento e decremento simetricos (Falha tambem) | nada |
| Mensagens de cascata diferenciadas | Ja existem via APS_MotivoEncaixe | nada |
| Nomes jaParado, mAPSProximoTick, APS_TICK_PROC | NAO existem nesta versao (usa mProx) | confirmar se o relatorio e de outra versao |
| EnableEvents/Calculation | Nao alterados em nenhum ponto do codigo atual | nada |

## Nao implementado
Atraso visual no card (precisa decidir armazenamento: coluna "Atraso"? atraso no fim ou no meio? entrada por celula ou formulario?).

## Testes pendentes (Excel)
Compilar; abrir/fechar o arquivo e conferir 1 Tick/min; colar >500 linhas e conferir encaixe; botao Atualizar sem pop-up; matriz A-G do relatorio mestre.
