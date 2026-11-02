# Relatório diário do grupo — Voz em Ação

Todo dia às 08h (Brasília) o GitHub gera o PDF da devolutiva do grupo de cada cliente e o Apps Script envia por e-mail
(vozemacao.drive@gmail.com, sabrinamelloalves78@gmail.com, martacristine@hotmail.com) — **somente se houver respostas novas**.

## 1. Planilha (Apps Script)
1. Abra a planilha → Extensões → Apps Script.
2. Substitua todo o código pelo conteúdo de `E:\CLAUDE\voz_em_acao_github_v22\apps_script_backend_v10.gs` → Salvar.
3. Engrenagem (Configurações do projeto) → Propriedades do script → Adicionar: `RELATORIO_KEY` = uma senha longa inventada (guarde-a).
4. Implantar → Gerenciar implantações → lápis → Versão: Nova versão → Implantar. (A URL não muda.)

## 2. GitHub
1. github.com → New repository → nome `voz-relatorios` → **Private** → Create.
2. Add file → Upload files: `gerar_relatorio.py`, `temas.json`, `logo.png` e a pasta `.github` (com `workflows/relatorio-diario.yml`).
3. Settings → Secrets and variables → Actions → New repository secret:
   - `APPS_SCRIPT_URL` = a URL que termina em /exec (a mesma do `config.js`)
   - `RELATORIO_KEY` = a mesma senha do passo 1.3
4. Aba Actions → habilite → "Relatório diário do grupo" → Run workflow (para testar; "1" no campo reenvia mesmo sem novidade).

## Variações do nome da empresa
No fim do `.gs`, a lista `REL_CLIENTES` define cada cliente com `apelidos` e `dominios` (e-mail). Também ignora acento,
maiúscula, espaços, Ltda/S.A./Brasil e erros pequenos de digitação. Cliente novo funciona sem cadastro (agrupado pelo nome);
para número de esperados e domínio, acrescente uma linha em `REL_CLIENTES`.
