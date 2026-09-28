# Mexico INPC Dashboard

Dashboard desenvolvido para acompanhamento dos releases de inflação do México.

A ferramenta consulta os dados oficiais do INEGI, atualiza as séries mensal e quinzenal e gera um HTML standalone com análise de headline, core, composição, momentum, sazonalidade, ranking e comparação com projeções pré-release.

## Arquivos

- `update_dashboard_final.py`: script principal de atualização
- `dashboard_template.html`: template do dashboard
- `btg_projections.xlsx`: input de projeções pré-release
- `requirements.txt`: dependências do projeto

## Execução

```bash
python update_dashboard_final.py --no-fallback
