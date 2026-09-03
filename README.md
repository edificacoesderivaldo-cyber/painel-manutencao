# 📊 Painel de Manutenção Predial — Com Google Sheets

Aplicativo Streamlit que **conecta diretamente ao seu Google Sheets** e gera dashboard em tempo real.

## ✨ O que você ganha

✅ **Sem upload manual** — dados sempre do Sheets  
✅ **Tempo real** — atualiza a cada 5 minutos  
✅ **Dashboard interativo** — KPIs, gráficos, filtros  
✅ **Funciona em qualquer navegador** — desktop, tablet, celular  
✅ **Seguro** — credenciais privadas (Google Cloud Service Account)  

---

## 🚀 Como usar

### **Pré-requisitos**

Você precisa ter feito os **Passos 1-5** descritos anteriormente:
- ✅ Criado um projeto no Google Cloud
- ✅ Ativado a Google Sheets API
- ✅ Criado uma Service Account
- ✅ Gerado o arquivo JSON das credenciais
- ✅ Compartilhado a planilha com a Service Account

Se não fez, volte aos passos anteriores!

---

### **Instalação (seu computador)**

1. **Baixe os arquivos:**
   - `painel_app_sheets.py`
   - `requirements_sheets.txt`
   - `credentials.json` (o arquivo que você baixou do Google Cloud)

2. **Coloque tudo na mesma pasta**

3. **Abra o terminal nessa pasta e execute:**

   ```bash
   pip install -r requirements_sheets.txt
   streamlit run painel_app_sheets.py
   ```

4. **Uma janela do navegador vai abrir automaticamente** com o painel conectado ao Sheets!

---

### **Hospedado na Nuvem (Streamlit Cloud - Gratuito)**

1. **Suba os arquivos para o GitHub:**
   - Crie um repositório novo
   - Coloque `painel_app_sheets.py` e `requirements_sheets.txt`
   - Coloque também o `credentials.json` (pode ser privado)

2. **No Streamlit Cloud:**
   - Clique "New App"
   - Selecione seu repositório
   - Aponte para `painel_app_sheets.py`
   - Clique Deploy

3. **Pronto!** Você tem um link público que atualiza sempre que você mexe no Sheets

---

## 📋 Estrutura do Painel

### 📈 **Indicadores Gerais**
- Total, Concluídos, Em execução, Paralisados, Aguardando orçamento

### 💰 **Saldo em Contrato**
- SESI e SENAI com barras de progresso

### 📊 **Status Geral**
- Gráfico pizza com distribuição

### 🏢 **Chamados por Unidade**
- Tabela e gráfico empilhado
- Valores investidos
- Cobertura de orçamentos

### 📋 **Lista de Chamados**
- Filtros por status, unidade, busca
- Detalhes completos de cada O.S.

---

## 🔄 Atualizações Automáticas

**Tudo que você mudar no Google Sheets, o painel atualiza automaticamente** (cache de 5 minutos).

Não precisa fazer upload, não precisa mexer em nada — é tudo automático!

---

## ⚙️ Configuração (se precisar mudar algo)

Se a planilha tiver um ID diferente, abra `painel_app_sheets.py` e mude essa linha:

```python
SHEET_ID = "1OWmc_hWsXznqeTMhFdfZCuZ9qqUjxk_1"  # ← Substitua pelo seu ID
```

---

## 🐛 Troubleshooting

**Erro: "credentials.json not found"**
- Coloque o arquivo JSON na mesma pasta do `painel_app_sheets.py`

**Erro: "Permission denied"**
- Verifique se você compartilhou a planilha com o email da Service Account

**Painel não atualiza**
- Aguarde 5 minutos (tempo do cache)
- Ou abra novamente a página do navegador

---

## 💬 Dúvidas?

É só me chamar — estou aqui para ajudar!

**Happy dashboarding! 📊**
