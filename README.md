# Desembolso — Backend API

Backend Flask conectado ao Google Sheets via gspread.

## Rodar localmente

```bash
pip install -r requirements.txt
# Coloque o arquivo JSON de credenciais como credentials.json na raiz
python app.py
```

## Deploy no Railway

1. Suba este repositório no GitHub (sem o credentials.json!)
2. Acesse railway.app → New Project → Deploy from GitHub
3. Adicione a variável de ambiente:
   - Nome: `GOOGLE_CREDENTIALS`
   - Valor: conteúdo completo do arquivo JSON de credenciais (cole tudo)
4. Deploy automático

## Endpoints

| Método | Rota | Descrição |
|--------|------|-----------|
| GET | /lancamentos | Lista todos os lançamentos |
| GET | /lancamentos?mes=05/2026 | Filtra por mês |
| GET | /lancamentos?categoria=MERCADO | Filtra por categoria |
| POST | /lancamentos | Adiciona lançamento |
| PUT | /lancamentos/<linha> | Atualiza lançamento |
| DELETE | /lancamentos/<linha> | Remove lançamento |
| GET | /orcamento | Retorna tabela do Desembolso Trimestral |
| GET | /resumo | Totais por categoria (pago) |
| GET | /health | Status do servidor |

## Variáveis de ambiente

- `GOOGLE_CREDENTIALS` — JSON completo da service account (obrigatório no deploy)
- `API_KEY` — chave simples para proteger a API pública. Quando definida, o app pede a chave no navegador e envia no header `X-API-Key`
- `CORS_ORIGINS` — lista opcional de origens permitidas, separadas por vírgula. Use o domínio público do app quando quiser restringir CORS
- `PORT` — porta do servidor (padrão: 5050)

## Segurança

⚠️ Nunca suba o arquivo `credentials.json` para o GitHub.
Para app publicado, defina uma `API_KEY` forte nas variáveis de ambiente do Render/Railway.
Adicione ao `.gitignore`:
```
credentials.json
*.json
__pycache__/
.env
```
