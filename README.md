# 🍺 Prime Bar — Integração YUZER → Google Sheets

App web que lê os relatórios exportados do YUZER e preenche automaticamente a planilha de fechamento de evento no Google Sheets.

---

## 📁 Estrutura do Projeto

```
primebar-app/
├── backend/
│   ├── app.py              # Servidor Flask (API)
│   └── requirements.txt    # Dependências Python
├── frontend/
│   └── index.html          # Interface visual
├── render.yaml             # Configuração de deploy (Render.com)
└── README.md
```

---

## ⚙️ Configuração — Credenciais do Google

Para o app conseguir escrever nas planilhas, você precisa criar uma "conta de serviço" no Google. Faça isso uma única vez:

### Passo 1 — Criar projeto no Google Cloud
1. Acesse [console.cloud.google.com](https://console.cloud.google.com)
2. Clique em **"Novo Projeto"** → dê o nome **"Prime Bar"** → Criar
3. No menu lateral: **APIs e Serviços** → **Biblioteca**
4. Pesquise **"Google Sheets API"** → Ativar

### Passo 2 — Criar conta de serviço
1. Menu lateral: **APIs e Serviços** → **Credenciais**
2. Clique em **"+ Criar Credenciais"** → **Conta de serviço**
3. Nome: `primebar-sheets` → Criar e continuar → Concluir
4. Clique na conta de serviço criada → aba **"Chaves"**
5. **Adicionar chave** → **Criar nova chave** → **JSON** → Baixar

### Passo 3 — Compartilhar planilhas com a conta de serviço
1. Abra o arquivo JSON baixado — copie o valor do campo `"client_email"` (algo como `primebar-sheets@...iam.gserviceaccount.com`)
2. Em CADA planilha de evento no Google Sheets:
   - Clique em **Compartilhar**
   - Cole o email da conta de serviço
   - Permissão: **Editor**
   - Salvar

---

## 🚀 Deploy no Render (gratuito)

### Passo 1 — Subir o código no GitHub
1. Crie uma conta em [github.com](https://github.com) se não tiver
2. Crie um repositório novo chamado `primebar-app`
3. Faça upload de todos os arquivos desta pasta

### Passo 2 — Criar app no Render
1. Acesse [render.com](https://render.com) → criar conta gratuita
2. Clique em **"New +"** → **Web Service**
3. Conecte sua conta do GitHub e selecione o repositório `primebar-app`
4. Configure:
   - **Build Command:** `pip install -r backend/requirements.txt`
   - **Start Command:** `cd backend && gunicorn app:app`
5. Em **Environment Variables**, adicione:
   - **Key:** `GOOGLE_CREDENTIALS`
   - **Value:** Cole TODO o conteúdo do arquivo JSON baixado no Passo 2 acima
6. Clique em **"Create Web Service"**

### Passo 3 — Servir o frontend
No mesmo repositório, configure também um **Static Site** no Render:
- **Publish directory:** `frontend`
- Aponte a variável de ambiente `API_URL` para a URL do seu web service

---

## 💻 Rodar Localmente (para testes)

```bash
# Instalar dependências
cd backend
pip install -r requirements.txt

# Configurar credenciais
export GOOGLE_CREDENTIALS='{"type":"service_account",...}'

# Iniciar servidor
python app.py

# Abrir frontend
# Abra o arquivo frontend/index.html no navegador
```

---

## 📊 O que o app preenche automaticamente

| Relatório YUZER | Aba da Planilha | Células |
|---|---|---|
| Produtos Vendidos | RELATORIO DE VENDA | Coluna B (Sistema) a partir da linha 5 |
| Exportação Caixas | FECHAMENTO CAIXAS | B3:H até último caixa |
| Painel de Vendas | RESUMO | B3:B7 (formas de pagamento) |

---

## 🔄 Fluxo de Uso (equipe administrativa)

1. Exportar 3 arquivos do YUZER após o evento
2. Acessar o link do app no navegador
3. Colar o link da planilha do evento
4. Fazer upload dos 3 arquivos
5. Conferir o preview dos dados
6. Clicar em **"Enviar para Google Sheets"**
7. ✅ Pronto — planilha preenchida automaticamente

---

## 📞 Suporte

Em caso de dúvidas sobre configuração, entre em contato com o administrador do sistema.
