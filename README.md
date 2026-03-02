# Monitoramento Power BI - Atualização Automática

Script que monitora status de refresh de datasets no Power BI/Fabric
e registra:

- Status na aba "Monitoramento Diário"
- Detalhes de erro na aba "Erros Dashboards"

## Requisitos

- Python 3.9+
- App Registration com permissão Power BI API
- Acesso ao Azure Blob Storage

## Instalação

```bash
python -m venv .venv
