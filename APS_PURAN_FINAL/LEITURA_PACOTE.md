# APS PURAN — PASTA FINAL DE ENTREGA

Caminho absoluto: `/tmp/kilo/APS_PURAN_FINAL`

## Conteudo

| Arquivo | Descricao |
|---|---|
| `modAPS_Base.bas` | Modulo corrigido (PEN-B-22, B-05, B-06, B-20, B-02, B-21, B-23 + correcao de APS_Ocupado em APS_ExcluirOperacao) |
| `modAPS_Motor.bas` | Modulo corrigido (PEN-B-18) |
| `modAPS_Cards.bas` | Modulo de referência (inalterado) |
| `modAPS_Calendario.bas` | Modulo de referência (inalterado) |
| `modAPS_SAP.bas` | Modulo de referência (inalterado) |
| `APS_Puran_TESTE_Fase_H.xlsm` | Copia de teste do original (onde aplicar as correcoes) |
| `APS_Puran_ORIGINAL.xlsm` | Copia de seguranca do original (NUNCA modifique) |
| `INSTRUCOES_IMPORTACAO.md` | Procedimento passo a passo para importar os modulos no Editor VBA |
| `LEITURA_PACOTE.md` | Este indice |

## Regras de uso

- Trabalhe SOMENTE em `APS_Puran_TESTE_Fase_H.xlsm`.
- Mantenha `APS_Puran_ORIGINAL.xlsm` fechado e intocado.
- NAO edite o arquivo binario do `.xlsm` por substitucao de texto.
- NAO altere velocidades, OEE, calendario, duracoes, quantidades, layout ou regras de produção.

## Estado do Git

- Branch: `main`
- Commit: `81b7362`
- Arquivos modificados (sem commit, sem push): `modAPS_Base.bas`, `modAPS_Motor.bas`