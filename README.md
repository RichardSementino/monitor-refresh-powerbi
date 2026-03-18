📊 Monitoramento Automatizado de Refresh Power BI (Fabric)

Automação desenvolvida para realizar o monitoramento inteligente e contínuo de atualizações (refresh) de datasets Power BI, registrando status operacional, histórico de execução e data oficial de atualização diretamente em uma base centralizada em Excel armazenada no Azure Blob Storage.

🎯 Objetivo do Projeto

Garantir visibilidade operacional sobre a atualização dos dashboards corporativos, permitindo:

Monitorar se os datasets atualizaram corretamente

Identificar falhas de refresh

Registrar histórico operacional

Controlar dashboards sem atualização recente

Apoiar governança e confiabilidade de dados

Reduzir dependência de monitoramento manual

🧠 Arquitetura da Solução
Power BI API (Fabric)
        ↓
Autenticação OAuth2 (Service Principal)
        ↓
Python Automation Script
        ↓
Consulta histórico de refresh datasets
        ↓
Atualização inteligente em arquivo Excel
        ↓
Upload automático no Azure Blob Storage
⚙️ Funcionalidades Implementadas
✅ Monitoramento de Refresh

Consulta automática da API do Power BI

Captura do status da última atualização

Captura da data e hora oficial do refresh (endTime)

✅ Atualização da Planilha de Monitoramento

Preenchimento automático da aba:

📄 Monitoramento Diário

Atualiza status OK / Erro

Atualiza apenas o dashboard correspondente

Baseado na data corrente

✅ Atualização do Catálogo de Reports

Preenchimento automático da aba:

📄 Catálogo_Reports

Atualiza:

Data da última atualização

Hora da última atualização

Informação baseada na API do Power BI (fonte oficial)

✅ Registro de Erros Operacionais

Preenchimento automático da aba:

📄 Erros Dashboards

Dataset

Status da execução

Detalhe técnico retornado pela API

Data e dia da semana

✅ Gestão Centralizada de Arquivo

Download único do Excel no início da execução

Atualizações em memória

Upload único ao final

Redução de custo e I/O

🔐 Autenticação

A automação utiliza Service Principal (App Registration) para autenticação segura via OAuth2.

Variáveis de ambiente necessárias:

AZURE_BLOB_CONNECTION_STRING
PBI_TENANT_ID
PBI_CLIENT_ID
PBI_CLIENT_SECRET
📦 Tecnologias Utilizadas

Python

Requests (Power BI REST API)

OpenPyXL

Azure Storage Blob SDK

OAuth2 Client Credentials Flow

Power BI / Microsoft Fabric

🚀 Roadmap do Projeto
🔄 Monitoramento Contínuo

Execução em janela operacional (ex: 24h)

Evolução para modelo orientado a eventos (refresh-trigger)

📊 Painel Operacional de Monitoramento

Dashboard Power BI dedicado

Visão de:

Últimos refreshs

Falhas

Dashboards sem atualização recente

⏳ Regra dos 90 dias

Identificação automática de dashboards sem atualização

Apoio à governança e limpeza de ativos analíticos

🤖 Inteligência Operacional (Futuro)

Classificação automática de erros

Sugestão de causas prováveis

Recomendações de resolução

Possível uso de GenAI no backend

💡 Aprendizados Técnicos

Integração Power BI API + Python

Gestão segura de credenciais

Automação de monitoramento operacional

Arquitetura orientada a eficiência de I/O

Governança de dados analíticos

Observabilidade de pipelines analíticos

🏁 Status do Projeto

🟢 Estrutura funcional implementada
🟡 Testes em ambiente corporativo dependem de liberação de Service Principal no Tenant
🔵 Evoluções arquiteturais planejadas