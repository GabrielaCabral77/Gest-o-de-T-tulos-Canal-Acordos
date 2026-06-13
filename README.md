# Canal Acordos — Cobrança & Prevenção

Portal interno de gestão de cobrança do Canal Acordos.

## Como fazer deploy na Vercel

### 1. Criar conta no GitHub (se não tiver)
Acesse https://github.com e crie uma conta gratuita.

### 2. Criar repositório
- Clique em **"New repository"**
- Nome: `canal-acordos`
- Deixe **Private** (recomendado para dados internos)
- Clique em **"Create repository"**

### 3. Subir os arquivos
Na página do repositório criado, clique em **"uploading an existing file"** e arraste toda a pasta do projeto (exceto `node_modules`).

Ou via terminal:
```bash
cd canal-acordos
git init
git add .
git commit -m "primeiro deploy"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/canal-acordos.git
git push -u origin main
```

### 4. Deploy na Vercel
- Acesse https://vercel.com e faça login com sua conta GitHub
- Clique em **"Add New Project"**
- Selecione o repositório `canal-acordos`
- Framework: **Vite** (detectado automaticamente)
- Clique em **"Deploy"**

Em ~2 minutos a URL estará disponível, ex: `canal-acordos.vercel.app`

### 5. Compartilhar com o time
Envie a URL gerada para o time. Para restringir acesso, na Vercel vá em **Settings → Password Protection** e defina uma senha.

---

## Atualização diária da base
No portal, clique em **📥 Atualizar Base** (canto superior direito) e carregue:
1. CSV de títulos exportado do sistema
2. XLSX do EMIT exportado do SAP

Os status, e-mails e histórico ficam salvos no navegador de cada usuário.

## Desenvolvimento local
```bash
npm install
npm run dev
```
Acesse http://localhost:5173
