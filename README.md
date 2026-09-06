📊 Painel de Manutenção Predial - SESI/SENAI
Sistema interativo para gestão e análise de chamados de manutenção predial das unidades SESI e SENAI.
🎯 Funcionalidades
✅ Upload de Planilha Excel - Carregue sua planilha com os dados de chamados
📈 KPIs Gerais - Visualize total, concluídos, em execução, paralisados, etc.
💰 Saldo em Contrato - Acompanhe utilização dos contratos SESI e SENAI
⏱️ Tempo Médio em Aberto - Análise de chamados por faixa de dias
📊 Distribuição de Status - Gráfico de rosca com distribuição de status
🏢 Chamados por Unidade - Gráfico empilhado + cards com valores investidos
💾 Download HTML - Exporte o dashboard como arquivo HTML interativo
🚀 Como Usar
Online (Streamlit Cloud)
Acesse: https://painel-manutencao-jgtxythyddzaqynntg7sfh.streamlit.app
Clique em "Envie seu arquivo Excel"
Selecione a planilha `CONTROLE_DE_O_S.xlsx`
O dashboard será gerado automaticamente
Clique em "Download HTML" para salvar localmente
Localmente (seu computador)
Pré-requisitos
Python 3.8+
pip (gerenciador de pacotes Python)
Instalação
Clone o repositório:
```bash
git clone https://github.com/edificacoesderivaldo-cyber/painel-manutencao.git
cd painel-manutencao
```
Instale as dependências:
```bash
pip install -r requirements.txt
```
Execute o app:
```bash
streamlit run painel_app_v2.py
```
Abra o navegador em: `http://localhost:8501`
📋 Formato da Planilha
A planilha deve ter as seguintes abas:
Aba "O.S" (header na linha 3)
Colunas obrigatórias:
NR - Número do registro
O.S - Número da ordem de serviço
UNIDADE - Nome da unidade (SESI/SENAI)
DATA DE ENVIO  DO CHAMADO - Data do chamado (dois espaços)
STATUS - Status do chamado (com espaço: "STATUS ")
DESCRIÇÃO DO SERVIÇO - Descrição breve
VALOR INICIAL DO SERVIÇO - Valor investido
Aba "SALDO" (header na linha 1)
Contém informações dos contratos SESI e SENAI
Linhas: contrato, saldo, etc.
🎨 Layout do Dashboard
Seção 1: KPIs Gerais
Total de Chamados
Concluídos
Em Execução
Paralisados
Abertos >30 dias
Valor Investido Total
Seção 2: Saldo em Contrato
SESI: Contrato, Utilizado, Barra de Progresso, Saldo Disponível
SENAI: Contrato, Utilizado, Barra de Progresso, Saldo Disponível
Seção 3: Tempo Médio em Aberto
Média de dias em aberto
Faixas: 0-30 dias (verde), 30-60 dias (amarelo), >60 dias (vermelho)
Seção 4: Distribuição de Status
Gráfico de rosca com proporções por status
Seção 5: Chamados por Unidade
Gráfico de barras horizontais empilhadas (cores por status)
Legenda de status
Cards com: Nome da unidade | Quantidade de chamados | Valor investido
📊 Status Disponíveis
Status	Cor
CONCLUIDO	🟢 Verde
EM EXECUÇÃO	🔵 Azul
PARALISADO	🔴 Vermelho
AGUARDANDO ORÇAMENTO	🟨 Amarelo
LIBERADO	🟩 Verde-água
PROJETO	⬜ Cinza
PLANEJAMENTO	🟪 Roxo
AGUARDANDO APROVAÇÃO	🟧 Laranja
🔧 Tecnologias
Frontend: HTML5 + Chart.js (gráficos)
Backend: Python + Streamlit (app web)
Data: Pandas (processamento de dados)
Excel: openpyxl (leitura de planilhas)
📝 Estrutura do Projeto
```
painel-manutencao/
├── painel_app_v2.py          # App Streamlit (NOVO)
├── requirements.txt           # Dependências Python
├── README.md                  # Este arquivo
└── .streamlit/
    └── config.toml           # Configurações Streamlit (opcional)
```
🐛 Troubleshooting
Erro: "ModuleNotFoundError: No module named 'streamlit'"
Solução: Instale as dependências com `pip install -r requirements.txt`
Erro: "Arquivo Excel não tem a aba 'O.S'"
Solução: Certifique-se de que o arquivo tem as abas corretas: "O.S" e "SALDO"
Erro: "Arquivo Excel não carrega"
Solução: Verifique se:
O arquivo é `.xlsx` ou `.xls`
A aba "O.S" tem headers na linha 3
As colunas estão com os nomes exatos (respeitando espaços)
📧 Contato
Desenvolvido para edificacoesderivaldo-cyber
Dúvidas ou sugestões? Abra uma issue no GitHub.
📄 Licença
Privado - Uso interno SESI/SENAI
---
Última atualização: 5 de setembro de 2026
