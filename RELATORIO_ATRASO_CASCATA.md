# APS PURAN - Atraso como tempo real de programacao + cascata nos dois sentidos

STATUS GERAL: codigo alterado por leitura/analise estatica. NAO compilado e NAO testado no Excel. Esta versao SUBSTITUI a versao "atraso so visual" (RELATORIO_CARDS_CASCATA.md, obsoleto). Inclui as correcoes de RELATORIO_ESTABILIDADE.md.

## Modelo do atraso
- Nova coluna **Atraso** em 02_Operacoes, logo depois de Setup (formato [h]:mm; criada automaticamente por APS_GarantirColunaAtraso em APS_Atualizar, mesmo mecanismo da coluna Setup; nenhuma outra coluna e movida). Digitar numero >= 1 = minutos.
- **Duracao** (coluna) = duracao ORIGINAL de producao (nunca sobrescrita). **Quantidade** nunca muda. **Fim** = termino programado = inicio + (original + atraso), pelo calendario/horarios de trabalho existentes.
- Em memoria (tOperacao) `dur` = original + atraso (tempo total programado) e `atraso` = parte adicional. Assim TODO o motor existente (encaixe, cascata, conflitos, status, cards) passa a considerar o tempo adicional sem reescrita. Gravacao: APS_DurOriginal(o) vai para Duracao; o.atraso vai para Atraso.
- Atraso fica no FIM da producao; Setup continua antes (Setup.fim = Producao.ini), sem alteracao.
- Sem linha, card ou operacao extra: a mesma linha e o mesmo card (card cresce porque o Fim aumenta; texto mostra "(atraso h:mm)").
- Persistencia: e a propria linha em 02_Operacoes (ID tecnico inalterado).

## Onde informar
1. Formulario (clique no card): campo "Atraso (hh:mm)" (vazio = sem atraso). Preview mostra Fim e cascata. Duracao exibida = original.
2. Direto na celula Atraso de 02_Operacoes (cascata automatica via APS_OpsAlterada/APS_ReencaixarLinha).

## Cascata
- Para a frente (inserir / aumentar atraso): APS_Encaixar com prioridade (operacao fica; seguintes avancam em cadeia, atomico, validado, com rollback) - codigo ja existente, agora recebendo o tempo total.
- Para tras (excluir / reduzir ou remover atraso / encurtar): correcao em APS_FecharVaos - o tempo que cada bloco puxado libera passa a contar como vao aberto por evento, entao TODAS as seguintes voltam (antes parava na primeira). Ociosidade planejada continua protegida.
- Editar o Fim a mao: a duracao total passa a ser ate o novo Fim; se o novo Fim deixar menos tempo que o atraso, a edicao e desfeita com mensagem (RestaurarEdicao tambem restaura Atraso/Duracao originais).

## Tabela
| Item | Modulo/rotina | Alteracao | Situacao |
|---|---|---|---|
| Atraso separado da duracao | Base: tOperacao.atraso, H_ATR, APS_DurOriginal, APS_LerOps, APS_GravarOp, APS_GravarHorario | coluna + campo | implementado, nao testado |
| Coluna Atraso | Base: APS_GarantirColunaAtraso; Cards: APS_Atualizar | cria se faltar | implementado, nao testado |
| Atraso ocupa maquina / move seguintes | Motor (sem mudanca de regra; usa dur total) | - | por analise |
| Cascata para tras em cadeia | Motor: APS_FecharVaos/APS_VaoLiberar | registra vao liberado | corrigido por leitura, nao testado |
| Edicao em celula (Atraso/Fim) | Base: APS_OpsAlterada; Motor: APS_ReencaixarLinha, RestaurarEdicao | tratamento | implementado, nao testado |
| Formulario | frmAPS_Operacao | campo Atraso, preview, duracao original | implementado, nao testado |
| Conflitos | DetectarConflitos (intersecao de intervalos, ja usa Fim total) | - | por analise |
| Setup/Limpeza | sem alteracao | - | preservado por analise |
| Dashboard/aba SAP | horas planejadas somam tempo total (inclui atraso); modAPS_SAP nao alterado (cria ops com atraso 0) | - | verificar |

## Riscos / pendencias
- Layout: 02_Operacoes ganha 1 coluna (como ocorreu com Setup). Formulas externas por posicao de coluna (se houver) devem ser conferidas; o codigo localiza colunas por cabecalho.
- Atraso conta em tempo de TRABALHO (calendario), nao tempo corrido. Se preferir corrido, e decisao a confirmar.
- Dashboard (horas) e coluna "Duracao" da aba de verificacao usam tempo total.
- Nada foi executado no Excel.

## Instalacao
1. Backup do .xlsm. 2. Importar/substituir modAPS_Motor, modAPS_Cards, modAPS_Base (modAPS_Calendario inalterado). 3. Colar frmAPS_Operacao_codigo.txt no codigo do formulario frmAPS_Operacao. 4. Depurar > Compilar VBAProject (enviar erros, se houver). 5. Abrir: a coluna Atraso e criada ao atualizar.
Retorno: restaurar o backup.

## Testes (Excel)
1. Op 4.000.000 comp/20h, atraso 2:00 -> Duracao 20:00, Atraso 2:00, Fim = ini + 22h de trabalho, quantidade igual, 1 linha, 1 card maior.
2. 4+ ops na mesma maquina; aumentar atraso da 1a -> demais avancam em cadeia sem sobreposicao; duracoes originais intactas.
3. Reduzir/remover atraso -> todas as seguintes voltam; intervalo ocioso planejado permanece.
4. Excluir op do meio -> seguintes voltam em cascata.
5. Inserir op em horario ocupado -> escolhida fica, seguintes deslocadas.
6. Op com Setup + atraso: Setup antes, atraso depois.
7. Atraso atravessando fim de expediente/limite do mes: respeita calendario; no limite do mes avisa e nada e gravado.
8. Salvar, fechar, reabrir: atraso mantido. 9. Sem linhas/cards duplicados.

---
# REVISAO FINAL (sobre a versao acima)
Metodo: releitura do codigo real + simulacao logica (simulacao_cascata_vao.py, calendario 24h, NAO e o VBA). Nada foi executado no Excel; compilacao NAO verificada.

| Item | Arquivo / rotina | Causa / achado | Correcao | Situacao | Teste |
|---|---|---|---|---|---|
| Cascata para tras parava na 1a operacao | Motor / APS_FecharVaos, APS_VaoLiberar | vao liberado pela op puxada nao era registrado, 2a op via vao "planejado" | registra o tempo liberado a cada bloco puxado; vetores crescem sob demanda | corrigido (por analise + simulacao) | simulacao: exclusao e reducao de atraso movem B, C, D (antes so B); ociosidade planejada intacta. Excel: pendente |
| Cascata para a frente com atraso | Motor / APS_Encaixar (dur total) | sem defeito encontrado; usa o tempo total | nenhuma | nao verificado (Excel) | roteiro nos testes |
| Atraso lido/gravado, duracao original preservada | Base / APS_LerOps, APS_GravarOp, APS_GravarHorario, APS_DurOriginal | dur em memoria = original + atraso; colunas gravam original e atraso separados; nenhum ponto soma duas vezes (revisados todos os escritores de .dur) | - | corrigido (por analise) | pendente |
| Contagem dupla do atraso | Form / previa e Validar | previa e salvamento usam o mesmo calculo (original + atraso uma vez) | - | corrigido (por analise) | pendente |
| Coluna Atraso criada/localizada | Base / APS_GarantirColunaAtraso; localizacao por cabecalho H_ATR | - | - | nao verificado (Excel) | pendente |
| Formatacao de colunas novas | Base / APS_CorrigirFormatoDados | verificava so colunas 1..10 (ignoraria Setup/Atraso) | agora APS_NColDados | corrigido (pequeno ajuste feito nesta revisao) | pendente |
| Edicao direta na celula Atraso / Fim | Base / APS_OpsAlterada; Motor / APS_ReencaixarLinha, RestaurarEdicao | Fim menor que o atraso invalidaria a duracao original | edicao desfeita com mensagem | corrigido (por analise) | pendente |
| Exclusao e atraso | Base / APS_ExcluirOperacao -> APS_LimparLinha | atraso fica na propria linha: sai junto (APS_NColDados) | nenhuma | corrigido (por analise) | pendente |
| Conflitos | Base / APS_DetectarConflitos | intersecao de intervalos usando o Fim total | nenhuma | nao verificado (Excel) | pendente |
| Setup/Limpeza | sem alteracao | - | - | preservado (por analise) | pendente |
| Persistencia | linha de 02_Operacoes (ID inalterado) | - | - | nao verificado | pendente |
| modAPS_SAP / ThisWorkbook / frmAPS_Config | nao alterados | SAP cria operacoes com atraso 0 | - | nao verificado | pendente |
| Compilacao VBA | todos | - | - | NAO verificado | pendente |

Arquivos finais completos (todos nesta pasta): modAPS_Base.bas, modAPS_Cards.bas, modAPS_Motor.bas, modAPS_Calendario.bas (inalterado), frmAPS_Operacao_codigo.txt, APS_modulos.zip. ThisWorkbook: sem alteracao.
Volta ao backup: fechar sem salvar, copiar de volta o .xlsm de backup (ou reimportar os .bas antigos exportados antes), reabrir. A coluna Atraso, se ja criada, pode ser apagada manualmente (o codigo antigo a ignora).
