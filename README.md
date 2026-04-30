# Relatório de Marca — Somos Valor

Benchmark competitivo da **Somos Valor** ([somosvalor.com.br](https://somosvalor.com.br/)) contra **Bankme**, **Finscale** e **Giro Tech**, comparando:

- Volume de busca de marca no Google (DataForSEO – Google Ads), janelas de março/2023 → março/2026
- Investimento estimado em **Meta Ads** (R$ 400 / anúncio) e **Google Ads** (R$ 200 / anúncio)
- Tráfego SEO orgânico mensal estimado
- Seguidores no Instagram

## Arquivos

| Arquivo | Descrição |
|---|---|
| `brand_report.py` | Gera o relatório (XLSX + HTML) consultando o DataForSEO |
| `relatorio_marca_somosvalor.html` | Relatório de apresentação (dark mode, pronto para abrir no navegador) |
| `relatorio_marca_somosvalor.xlsx` | Mesmos dados em planilha |
| `construtoras.txt` / `construtora1.txt` | Lista paralela das 20 maiores construtoras do Brasil (ranking INTEC 2026) |
| `100maioresconstrutoras.com.br-0.md` | Fonte usada para o ranking de construtoras |

## Rodar

```bash
python brand_report.py
```

Requer Python 3 com `openpyxl`. As credenciais do DataForSEO estão no client embutido na skill local `pesquisa-marca-dataforseo`.

## Estrutura do relatório HTML

1. **Cabeçalho** com KPIs da Somos Valor
2. **Tabela comparativa de marca** — crescimento 1 ano / 3 anos + volumes 03/2026 → 03/2023
3. **Comparativo por canal** — Meta Ads, Google Ads, SEO Orgânico, Instagram (cada bloco ordenado do maior para o menor; Meta e Google Ads incluem coluna de Estimativa de Investimento)

A linha da Somos Valor fica destacada em todos os blocos.
