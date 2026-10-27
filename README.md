# UEA nas Redes CAPES Global — mapeamento das 3 redes

Dashboard autocontido (`index.html`) com a participação da UEA nas redes CAPES Global:
**Rede Conecta Mundi (UNESP)**, **Rede INTER-NÓS (UFMG)** e **SustentNET (UNICAMP)**.

- **Visão geral** — as 3 redes, PPGs da UEA em cada rede com valores previstos (set/2027) e perfil de notas CAPES.
- **Uma aba por rede** — mapa das IES, perfil de cada IES, matriz área temática × universidade,
  cards das áreas (PPGs de todas as IES) e conferência comitê UEA × Anexo I.
- **Explorar PPGs** — as 738 vinculações dos Anexos I, com filtros, ordenação e exportação CSV.

## Gerar novamente

```bash
python3 build_dashboard.py   # requer openpyxl e pdftotext (poppler)
```

Fontes: `CAPES_Global_Mapeamento_3_Redes_PPGs_IES.xlsm` (Anexos I das propostas),
`fontes/Balanco_CAPES_Global_2027…pdf` (Setembro 2027 | Bolsas a atender) e o cadastro do
comitê CAPES Global do Gerenciador PROPESP (caminho local em `build_dashboard.py`).
