# Controle de Férias N2 · 2027

Painel: https://iatimcll-blip.github.io/ferias-n2/

## Conteúdo

| Caminho | O que é |
|---|---|
| `Painel Férias N2 2027.html` | Painel completo em um único arquivo: dados, gráficos, anti-colisão, edição de datas, hierarquia e a Vivi (assistente de inconsistências) |
| `index.html` | Redireciona o GitHub Pages para o painel |
| `dados/estado.json` | Alterações feitas no painel e compartilhadas entre aparelhos. É criado pelo próprio painel; não edite à mão |
| `dados/base.xlsx` | Última planilha de férias carregada no painel. É usada pelos outros aparelhos e pelo relatório |
| `Controle_Ferias_N2_2027.xlsx` | Planilha de controle de férias (base do painel) |
| `HIERARQUIA/HIERARQUIA.xlsx` | Hierarquia GO › GA › colaborador |
| `scripts/anti_colisao.py` | Anti-colisão direto na planilha (mesmas regras do painel) |
| `.claude/agents/ajuste-ferias.md` | Agente do Claude Code que roda o anti-colisão e explica o resultado |
| `ferias_regras.json` | Limites usados pelo script |
| `relatorios/` | Saídas do script: alterações (.csv), resumo (.json) e planilha ajustada |

## Sincronização entre aparelhos

O que muda no painel fica salvo em `dados/estado.json` neste repositório: edições de datas, cadeados, regras e proposta aplicada do anti-colisão, hierarquia carregada e **base de férias carregada** (a planilha vai para `dados/base.xlsx`). Todo aparelho que abre o painel carrega a versão mais recente e confere de novo a cada minuto e ao voltar para a aba.

- **Só visualizar:** não precisa de nada. Qualquer aparelho recebe as alterações e a base mais nova.
- **Editar e sincronizar:** abra uma vez o **link da equipe**. O aparelho passa a enviar as alterações sozinho, sem configurar nada.
- **Gerar o link da equipe** (uma vez): crie um token *fine-grained* do GitHub com acesso só a este repositório e permissão **Contents: Read and write**. Cole o token em **Sincronização**, no topo do painel, e clique em **Copiar link da equipe**. O link funciona como uma senha: envie só para a equipe. Para cortar o acesso, revogue o token no GitHub e gere um link novo.
- Se dois aparelhos alteram ao mesmo tempo, o painel avisa o conflito e você escolhe qual versão fica.

## Atualizar a base

- **Nova planilha de férias:** no painel, use **Carregar base .xlsx** num aparelho habilitado pelo link da equipe. A planilha vai para todos os aparelhos. A base embutida no HTML só é usada enquanto nenhuma base foi sincronizada.
- **Nova hierarquia:** na aba Hierarquia, use **Carregar hierarquia .xlsx** (colunas GO, GA, NOME). Ela é sincronizada entre os aparelhos.
- **Levar as datas para a planilha:** use **Gerar relatório .xlsx**.

## Script

```
python -I scripts/anti_colisao.py --simular     # só calcula
python -I scripts/anti_colisao.py               # gera a planilha ajustada em relatorios/
```

> Repositório público: as planilhas e o painel contêm nomes e datas de férias de colaboradores.
