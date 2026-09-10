# DiscoverON — análise financeira da abertura da escola

[![site](https://img.shields.io/badge/site-analise--financeira-2ea44f)](https://analise-financeira-production.up.railway.app)
Deploy automático: todo push na `main` republica o site.

Modelagem financeira para a abertura de uma franquia DiscoverON (inglês + robótica):
cenários pessimista / moderado / otimista, quadro mês a mês, simulador interativo,
canais de aquisição com CAC e o controle dos custos de implantação.

## O que tem aqui

```
site/index.html   página publicada (HTML+CSS+JS, sem build, sem dependência externa)
Dockerfile        nginx:alpine servindo o site/index.html na porta 80
planilha/build_xlsx.py                      gerador da planilha (openpyxl)
planilha/DiscoverON-projecao-corrigida.xlsx planilha gerada
```

A página é um arquivo único e autocontido. Basta abrir o `site/index.html` no navegador
para rodar local — não precisa de servidor nem de instalar nada.

## Publicação (Railway)

Projeto `discoveron-ponto-tiburcio-511` · ambiente `production` · serviço `analise-financeira`
→ https://analise-financeira-production.up.railway.app

O serviço está conectado a este repositório: **todo push na `main` publica automaticamente.**
Não é preciso rodar nada — edite o `site/index.html`, commite e faça o push.

```bash
git add -A && git commit -m "..." && git push
```

O build usa o `Dockerfile` da raiz (nginx na porta 80). Se ele sumir, o Railway cai no
Caddy na porta 8080, o domínio aponta para a 80 e o site responde 502.

Para acompanhar ou forçar um deploy:

```bash
railway link --project c6100e3f-eb91-4570-a463-2652573e3947 --environment production
railway service link analise-financeira
railway deployment list      # status dos últimos deploys
railway logs                 # logs de build e runtime
railway redeploy             # republica o último deploy
```

## Regenerar a planilha

```bash
python3 -m venv venv && ./venv/bin/pip install openpyxl
./venv/bin/python planilha/build_xlsx.py
```

## Premissas do modelo

| Item | Valor |
|---|---|
| Matrícula | R$ 200 |
| Mensalidade | R$ 280 |
| Material didático | custo R$ 375 em 12x · venda R$ 900 (≈ R$ 810 antecipado no cartão) |
| Churn | 10% ao mês |
| Inadimplência | 10% |
| Aluguel | R$ 12.000 |
| Impostos | Simples Nacional Anexo III (progressivo, com RBT12 proporcionalizado) |
| Salas | 4 de inglês (8 alunos · 2h/semana) + 1 de informática (10 alunos · 1h/semana) |
| Horário | 9h–20h, exceto 12h–13h, com coeficiente de ocupação |
| CAPEX | R$ 198.713 (estimativas + 15% de reserva) |

Set/26 e Out/26 são meses de obra (sem custo operacional); a escola abre em Dez/26 (M1),
com horizonte até M16.

## Estado dos dados

Os lançamentos do playground, dos custos mês a mês, dos canais de aquisição e dos custos
de implantação ficam no `localStorage` do navegador — **não** são versionados aqui. Use os
botões de Exportar/Importar de cada seção para levar os dados para outra máquina.
