# Controle de Férias N2 · 2027

Painel: https://iatimcll-blip.github.io/ferias-n2/

## Conteúdo

| Caminho | O que é |
|---|---|
| `Painel Férias N2 2027.html` | Painel completo em um único arquivo: dados, gráficos, anti-colisão, edição de datas, hierarquia e a Vivi (assistente de inconsistências) |
| `index.html` | Redireciona o GitHub Pages para o painel |
| `dados/estado.json` | Alterações feitas no painel e compartilhadas entre aparelhos. É criado pelo próprio painel; não edite à mão |
| `Controle_Ferias_N2_2027.xlsx` | Planilha de controle de férias (base do painel) |
| `HIERARQUIA/HIERARQUIA.xlsx` | Hierarquia GO › GA › colaborador |
| `scripts/anti_colisao.py` | Anti-colisão direto na planilha (mesmas regras do painel) |
| `.claude/agents/ajuste-ferias.md` | Agente do Claude Code que roda o anti-colisão e explica o resultado |
| `ferias_regras.json` | Limites usados pelo script |
| `relatorios/` | Saídas do script: alterações (.csv), resumo (.json) e planilha ajustada |

## Sincronização entre aparelhos

O que muda no painel fica salvo em `dados/estado.json` neste repositório: edições de datas, cadeados, regras e proposta aplicada do anti-colisão, e hierarquia carregada. Todo aparelho que abre o link carrega a versão mais recente e confere de novo a cada 2 minutos.

- **Só visualizar:** não precisa de nada.
- **Editar e sincronizar:** em **Sincronização** (no topo do painel), cole um token *fine-grained* do GitHub com acesso só a este repositório e permissão **Contents: Read and write**. O token fica guardado apenas naquele navegador.
- Se dois aparelhos alteram ao mesmo tempo, o painel avisa o conflito e você escolhe qual versão fica.

## Atualizar a base

- **Nova planilha de férias:** no painel, use **Carregar base .xlsx**. Isso vale só para aquele navegador. Para todos verem, a base embutida no HTML precisa ser atualizada.
- **Nova hierarquia:** na aba Hierarquia, use **Carregar hierarquia .xlsx** (colunas GO, GA, NOME). Ela é sincronizada entre os aparelhos.
- **Levar as datas para a planilha:** use **Gerar relatório .xlsx**.

## Script

```
python -I scripts/anti_colisao.py --simular     # só calcula
python -I scripts/anti_colisao.py               # gera a planilha ajustada em relatorios/
```

> Repositório público: as planilhas e o painel contêm nomes e datas de férias de colaboradores.
