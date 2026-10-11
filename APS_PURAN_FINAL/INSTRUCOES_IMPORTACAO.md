# INSTRUCOES DE IMPORTACAO — APS Puran

## Preparacao
1. Abra `APS_Puran_TESTE_Fase_H.xlsm` no Excel.
2. Mantenha `APS_Puran_ORIGINAL.xlsm` fechado e intocado.
3. Trabalhe SOMENTE na copia de teste.

## Passo a passo
1. Pressione **Alt + F11** para abrir o Editor VBA.
2. No Project Explorer (lado direito), clique com o botao direito no projeto.
3. Escolha **Importar** (Import File).
4. Selecione `modAPS_Base.bas` e clique **Abrir**. O modulo `modAPS_Base` existente sera substituido.
5. Repita o passo 4 para `modAPS_Motor.bas`.
6. Verifique no Project Explorer que apenas os modulos esperados estao presentes:
   - `modAPS_Base`
   - `modAPS_Motor`
   - `modAPS_Cards`
   - `modAPS_Calendario`
   - `modAPS_SAP`
   - `EstaPastaDeTrabalho` (ThisWorkbook)
   - `frmAPS_Config`
   - `frmAPS_Operacao`
7. Salve o arquivo (Ctrl + S). O Excel deve sugerir o formato `.xlsm` — confirme.

## Compilar
1. No Editor VBA, clique em **Depurar** (Debug).
2. Escolha **Compilar VBAProject** (Compile VBAProject).
3. Anote o resultado:
   - Se houver erro, registre a mensagem, o nome do modulo e o numero de linha.
   - Corrija somente o defeito comprovado.
   - Compile novamente ate nao haver erros.

## Antes de executar qualquer macro
- Feche e reabra o arquivo para garantir que os novos modulos foram carregados.
- Confirme que a aba `02_Operacoes` e `1_Planejamento` estao presentes.
- Confirme que a coluna "OP" (depois de Produto) e a coluna "Setup" (depois de Duracao) existem.

## Procedimento de recuperacao
Se a importacao ou a compilacao falharem:
1. Feche o arquivo de teste.
2. Exclua `APS_Puran_TESTE_Fase_H.xlsm`.
3. Copie `APS_Puran_ORIGINAL.xlsm` para `APS_Puran_TESTE_Fase_H.xlsm`.
4. Reinicie pelo passo 1 deste arquivo.

## Limitacoes conhecidas
- Neste ambiente nao ha Excel. Estas instrucoes sao para execucao manual em uma estacao com Microsoft Excel.
- A compilacao nao foi realizada. Qualquer erro de sintaxe sera descoberto apenas na etapa "Compilar".
- Nenhuma macro foi executada. A matriz de testes esta inicialmente em **NAO EXECUTADO**.

## Matriz de testes (10 cenarios)

| # | Cenario | Status |
|---|---|---|
| 1 | Migração de Setup com falha intermediaria e confirmacao do rollback | NAO EXECUTADO |
| 2 | Limpeza de Setups orfaos e restauracao correta de APS_Ocupado | NAO EXECUTADO |
| 3 | Sincronizacao com mudancas de horario e dados, incluindo falha e rollback | NAO EXECUTADO |
| 4 | Agendamento com cancelamento anterior bem-sucedido e malsucedido, verificando duplicacao e recuperacao | NAO EXECUTADO |
| 5 | Falha de persistencia em APS_NovoID, verificando retorna "" e nao disponibiliza um ID nao persistido | NAO EXECUTADO |
| 6 | Exclusao de operacao, verificando que IDs excluidos nao sao reutilizados | NAO EXECUTADO |
| 7 | Integridade dos vinculos entre Producao, Setup e Limpeza | NAO EXECUTADO |
| 8 | Aba 02_Operacoes: uma unica linha por operacao de producao e Setup apresentado somente na coluna correspondente, sem linha duplicada de Setup | NAO EXECUTADO |
| 9 | Verificacao de que alteracoes de status nao deslocam as operacoes | NAO EXECUTADO |
| 10 | Conferencia do calendario, OEE, velocidades, duracoes, quantidades em comprimidos e demais parametros preexistentes | NAO EXECUTADO |