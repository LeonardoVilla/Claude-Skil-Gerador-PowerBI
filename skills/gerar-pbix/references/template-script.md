# Template de Script Python — Gerador de .pbix

Script completo e validado para gerar arquivos `.pbix` do Power BI. Copie e adapte para cada projeto.

## Dependências

Apenas biblioteca padrão do Python: `json`, `zipfile`, `os`, `copy`, `zlib`, `random`, `string`.

---

## Estrutura completa do script

```python
"""
Gerador de .pbix — <Nome do Projeto>
Template: <caminho do .pbix template>
Fonte: <caminho do .xlsx ou banco>
"""

import json, zipfile, os, copy, zlib, random, string

# ─── Utilitários ─────────────────────────────────────────────────────────────

def uid():
    return ''.join(random.choices(string.hexdigits[:16], k=20))

# OBRIGATÓRIO: alias fixos por tabela — NUNCA usar table[0].lower()
# Ajuste para as tabelas do seu projeto
TABLE_ALIAS = {
    "NOME_TABELA_1": "t1",
    "NOME_TABELA_2": "t2",
    # tabelas com espaços/acentos precisam de alias explícito:
    # "TABELA COM ESPAÇO ": "ts",  # inclua o espaço final se ele existir no nome
}

def get_src(table):
    return TABLE_ALIAS.get(table, table[:2].lower().replace(' ', '_'))

def pos(x, y, w, h, z=0, tab=0):
    return {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": tab}

def vc(x, y, w, h, cfg, z=0, tab=0):
    return {
        "x": x, "y": y, "z": z, "width": w, "height": h,
        "config": json.dumps(cfg, ensure_ascii=False, separators=(',', ':')),
        "filters": "[]", "tabOrder": tab
    }

# ─── Visuais ─────────────────────────────────────────────────────────────────

def shape_vc(x, y, w, h, color="#15314F", z=0, tab=0):
    """Retângulo colorido de fundo."""
    name = uid()
    cfg = {
        "name": name,
        "layouts": [{"id": 0, "position": pos(x, y, w, h, z, tab)}],
        "singleVisual": {
            "visualType": "shape",
            "drillFilterOtherVisuals": True,
            "objects": {
                "line": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                "fill": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "fillColor": {"solid": {"color": {"expr": {"Literal": {"Value": f"'{color}'"}}}}},
                    "transparency": {"expr": {"Literal": {"Value": "0D"}}}
                }}],
                "rotation": [{"properties": {"angle": {"expr": {"Literal": {"Value": "0D"}}}}}]
            },
            "vcObjects": {"background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]}
        }
    }
    return vc(x, y, w, h, cfg, z, tab)

def textbox_vc(x, y, w, h, text, size="10pt", color="#ffffff", bold=False, z=0, tab=0, pad_top=4, pad_left=8):
    """Caixa de texto estático."""
    name = uid()
    style = {"fontSize": size, "color": color}
    if bold:
        style["fontWeight"] = "bold"
    cfg = {
        "name": name,
        "layouts": [{"id": 0, "position": pos(x, y, w, h, z, tab)}],
        "singleVisual": {
            "visualType": "textbox",
            "drillFilterOtherVisuals": True,
            "objects": {"general": [{"properties": {"paragraphs": [
                {"textRuns": [{"value": text, "textStyle": style}]}
            ]}}]},
            "vcObjects": {
                "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                "padding": [{"properties": {
                    "top": {"expr": {"Literal": {"Value": f"{pad_top}D"}}},
                    "left": {"expr": {"Literal": {"Value": f"{pad_left}D"}}}
                }}]
            }
        }
    }
    return vc(x, y, w, h, cfg, z, tab)

def card_vc(x, y, w, h, table, field, label, agg="CountNonNull", z=0, tab=0):
    """
    Card KPI. Fundo cinza claro, número escuro — legível em qualquer fundo.
    agg: "CountNonNull" | "Count" | "Sum" | "Max" | "Min"
    """
    name = uid()
    src = get_src(table)
    agg_map = {"Count": 3, "CountNonNull": 5, "Sum": 0, "Max": 4, "Min": 2}
    fn = agg_map.get(agg, 5)
    ref = f"{agg}({table}.{field})"
    cfg = {
        "name": name,
        "layouts": [{"id": 0, "position": pos(x, y, w, h, z, tab)}],
        "singleVisual": {
            "visualType": "card",
            "projections": {"Values": [{"queryRef": ref}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": src, "Entity": table, "Type": 0}],
                "Select": [{"Aggregation": {
                    "Expression": {"Column": {
                        "Expression": {"SourceRef": {"Source": src}},
                        "Property": field
                    }},
                    "Function": fn
                }, "Name": ref, "NativeReferenceName": label}]
            },
            "columnProperties": {ref: {"displayName": label}},
            "drillFilterOtherVisuals": True,
            "hasDefaultSort": True,
            "objects": {
                "categoryLabels": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "fontSize": {"expr": {"Literal": {"Value": "9D"}}},
                    "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#555555'"}}}}}
                }}],
                "labels": [{"properties": {
                    "fontSize": {"expr": {"Literal": {"Value": "28D"}}},
                    "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#15314F'"}}}}}
                }}]
            },
            "vcObjects": {
                "background": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#F2F2F2'"}}}}},
                    "transparency": {"expr": {"Literal": {"Value": "0D"}}}
                }}],
                "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
                "title": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
            }
        }
    }
    return vc(x, y, w, h, cfg, z, tab)

def donut_vc(x, y, w, h, table, cat_field, val_field, title="", agg="CountNonNull", z=0, tab=0):
    """Gráfico de rosca. cat_field = categoria (eixo), val_field = valor agregado."""
    name = uid()
    src = get_src(table)
    agg_map = {"Count": 3, "CountNonNull": 5, "Sum": 0}
    fn = agg_map.get(agg, 5)
    val_ref = f"{agg}({table}.{val_field})"
    cat_ref = f"{table}.{cat_field}"
    cfg = {
        "name": name,
        "layouts": [{"id": 0, "position": pos(x, y, w, h, z, tab)}],
        "singleVisual": {
            "visualType": "donutChart",
            "projections": {
                "Y": [{"queryRef": val_ref}],
                "Category": [{"queryRef": cat_ref, "active": True}]
            },
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": src, "Entity": table, "Type": 0}],
                "Select": [
                    {"Aggregation": {
                        "Expression": {"Column": {
                            "Expression": {"SourceRef": {"Source": src}},
                            "Property": val_field
                        }},
                        "Function": fn
                    }, "Name": val_ref, "NativeReferenceName": f"Contagem de {val_field}"},
                    {"Column": {
                        "Expression": {"SourceRef": {"Source": src}},
                        "Property": cat_field
                    }, "Name": cat_ref, "NativeReferenceName": cat_field}
                ]
            },
            "drillFilterOtherVisuals": True,
            "hasDefaultSort": True,
            "objects": {
                "legend": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "position": {"expr": {"Literal": {"Value": "'RightCenter'"}}}
                }}],
                "labels": [{"properties": {
                    "labelStyle": {"expr": {"Literal": {"Value": "'Percent of total'"}}},
                    "percentageLabelPrecision": {"expr": {"Literal": {"Value": "0L"}}}
                }}]
            },
            "vcObjects": {
                "title": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                    "fontColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#15314F'"}}}}}
                }}],
                "background": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}
                }}],
                "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
            }
        }
    }
    return vc(x, y, w, h, cfg, z, tab)

def bar_vc(x, y, w, h, table, cat_field, val_field, title="", agg="CountNonNull",
           orientation="horizontal", z=0, tab=0):
    """
    Gráfico de barras ou colunas.
    orientation: "horizontal" → barChart | "vertical" → columnChart
    """
    name = uid()
    src = get_src(table)
    agg_map = {"Count": 3, "CountNonNull": 5, "Sum": 0}
    fn = agg_map.get(agg, 5)
    val_ref = f"{agg}({table}.{val_field})"
    cat_ref = f"{table}.{cat_field}"
    vtype = "barChart" if orientation == "horizontal" else "columnChart"
    cfg = {
        "name": name,
        "layouts": [{"id": 0, "position": pos(x, y, w, h, z, tab)}],
        "singleVisual": {
            "visualType": vtype,
            "projections": {
                "Y": [{"queryRef": val_ref}],
                "Category": [{"queryRef": cat_ref, "active": True}]
            },
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": src, "Entity": table, "Type": 0}],
                "Select": [
                    {"Aggregation": {
                        "Expression": {"Column": {
                            "Expression": {"SourceRef": {"Source": src}},
                            "Property": val_field
                        }},
                        "Function": fn
                    }, "Name": val_ref, "NativeReferenceName": f"Contagem de {val_field}"},
                    {"Column": {
                        "Expression": {"SourceRef": {"Source": src}},
                        "Property": cat_field
                    }, "Name": cat_ref, "NativeReferenceName": cat_field}
                ]
            },
            "drillFilterOtherVisuals": True,
            "hasDefaultSort": True,
            "objects": {
                "labels": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}}],
                "categoryAxis": [{"properties": {
                    "showAxisTitle": {"expr": {"Literal": {"Value": "false"}}}
                }}],
                "valueAxis": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
            },
            "vcObjects": {
                "title": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                    "fontColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#15314F'"}}}}}
                }}],
                "background": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}
                }}],
                "border": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
            }
        }
    }
    return vc(x, y, w, h, cfg, z, tab)

def slicer_vc(x, y, w, h, table, field, z=0, tab=0):
    """Filtro dropdown para o usuário selecionar valores."""
    name = uid()
    src = get_src(table)
    col_ref = f"{table}.{field}"
    cfg = {
        "name": name,
        "layouts": [{"id": 0, "position": pos(x, y, w, h, z, tab)}],
        "singleVisual": {
            "visualType": "slicer",
            "projections": {"Values": [{"queryRef": col_ref, "active": True}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": src, "Entity": table, "Type": 0}],
                "Select": [{"Column": {
                    "Expression": {"SourceRef": {"Source": src}},
                    "Property": field
                }, "Name": col_ref, "NativeReferenceName": field}]
            },
            "drillFilterOtherVisuals": True,
            "hasDefaultSort": True,
            "objects": {
                "data": [{"properties": {"mode": {"expr": {"Literal": {"Value": "'Dropdown'"}}}}}],
                "selection": [{"properties": {
                    "selectAllCheckboxEnabled": {"expr": {"Literal": {"Value": "false"}}},
                    "singleSelect": {"expr": {"Literal": {"Value": "false"}}}
                }}],
                "header": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "fontColor": {"solid": {"color": {"expr": {"Literal": {"Value": "'#FFFFFF'"}}}}}
                }}]
            },
            "vcObjects": {
                "background": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
            }
        }
    }
    return vc(x, y, w, h, cfg, z, tab)

# ─── Página exemplo ───────────────────────────────────────────────────────────

def page_exemplo():
    """Adapte esta função para cada página do seu painel."""
    vcs = []
    t = [0]
    def nt():
        t[0] += 1000
        return t[0]

    # Header azul escuro
    vcs.append(shape_vc(0, 0, 1280, 66, "#15314F", z=0, tab=nt()))
    vcs.append(textbox_vc(0, 4, 860, 36, "Nome do Painel",
                          "17pt", "#FFFFFF", True, z=100, tab=nt(), pad_top=6, pad_left=24))
    vcs.append(textbox_vc(0, 42, 700, 22, "Subtítulo ou filtro ativo",
                          "9pt", "#A9C0D8", False, z=100, tab=nt(), pad_top=2, pad_left=24))

    # Faixa de slicers
    vcs.append(shape_vc(0, 66, 1280, 42, "#1B3F60", z=200, tab=nt()))
    vcs.append(slicer_vc(10, 70, 230, 34, "NOME_TABELA", "CAMPO_FILTRO", z=500, tab=nt()))

    # Cards KPI (sem shape por baixo — o card tem fundo próprio)
    vcs.append(card_vc(10, 116, 295, 90, "NOME_TABELA", "CAMPO_ID",
                       "Label do Card", agg="CountNonNull", z=300, tab=nt()))

    # Gráfico de rosca
    vcs.append(donut_vc(10, 216, 390, 310,
                        "NOME_TABELA", "CAMPO_CATEGORIA", "CAMPO_VALOR",
                        title="Título do Donut", agg="CountNonNull", z=300, tab=nt()))

    # Gráfico de barras
    vcs.append(bar_vc(410, 216, 860, 310,
                      "NOME_TABELA", "CAMPO_CATEGORIA", "CAMPO_VALOR",
                      title="Título das Barras", agg="CountNonNull",
                      orientation="horizontal", z=300, tab=nt()))

    page_cfg = json.dumps({"objects": {"background": [{"properties": {
        "transparency": {"expr": {"Literal": {"Value": "0D"}}},
        "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#EEF1F6'"}}}}}
    }}]}}, ensure_ascii=False, separators=(',', ':'))

    return {
        "id": 0, "name": uid(), "displayName": "Nome da Pagina",
        "filters": "[]", "ordinal": 0, "width": 1280, "height": 720,
        "displayOption": 1, "visualContainers": vcs, "config": page_cfg
    }

# ─── Layout raiz ─────────────────────────────────────────────────────────────

def build_layout(pages):
    root_cfg = json.dumps({
        "version": "5.73",
        "themeCollection": {"baseTheme": {
            "name": "CY26SU05", "type": 2,
            "version": {"visual": "2.9.0", "report": "3.3.0", "page": "2.3.1"}
        }},
        "activeSectionIndex": 0,
        "defaultDrillFilterOtherVisuals": True,
        "linguisticSchemaSyncVersion": 2,
        "settings": {
            "useNewFilterPaneExperience": True,
            "allowChangeFilterTypes": True,
            "useStylableVisualContainerHeader": True,
            "queryLimitOption": 6,
            "useEnhancedTooltips": True,
            "exportDataMode": 1,
            "useDefaultAggregateDisplayName": True
        },
        "objects": {
            "section": [{"properties": {"verticalAlignment": {"expr": {"Literal": {"Value": "'Top'"}}}}}],
            "outspacePane": [{"properties": {"expanded": {"expr": {"Literal": {"Value": "false"}}}}}]
        }
    }, ensure_ascii=False, separators=(',', ':'))

    return {
        "id": 0,
        "resourcePackages": [{"resourcePackage": {
            "name": "SharedResources", "type": 2,
            "items": [{"type": 202, "path": "BaseThemes/CY26SU05.json", "name": "CY26SU05"}],
            "disabled": False
        }}],
        "sections": pages,
        "config": root_cfg,
        "layoutOptimization": 0
    }

# ─── Empacotar — NÃO alterar esta função ─────────────────────────────────────

def pack_pbix(source_pbix, output_pbix, layout_json):
    """
    Substitui o Report/Layout e zera o SecurityBindings.
    NUNCA alterar a lógica de SecurityBindings ou compress_type.
    """
    layout_str = json.dumps(layout_json, ensure_ascii=False, separators=(',', ':'))
    layout_bytes = layout_str.encode('utf-16-le')  # UTF-16 LE sem BOM — obrigatório

    if os.path.exists(output_pbix):
        os.remove(output_pbix)

    with zipfile.ZipFile(source_pbix, 'r') as zin:
        orig_info = next(i for i in zin.infolist() if i.filename == 'Report/Layout')
        new_info = copy.copy(orig_info)
        new_info.file_size = len(layout_bytes)
        new_info.CRC = zlib.crc32(layout_bytes) & 0xFFFFFFFF

        with zipfile.ZipFile(output_pbix, 'w', allowZip64=True) as zout:
            for item in zin.infolist():
                if item.filename == 'Report/Layout':
                    zout.writestr(new_info, layout_bytes, compress_type=orig_info.compress_type)
                elif item.filename == 'SecurityBindings':
                    # CRITICO: zerar DPAPI blob — sem isso Power BI rejeita com MashupValidationError
                    sb = copy.copy(item)
                    sb.file_size = 0
                    sb.CRC = zlib.crc32(b'') & 0xFFFFFFFF
                    zout.writestr(sb, b'', compress_type=item.compress_type)
                else:
                    zout.writestr(copy.copy(item), zin.read(item.filename),
                                  compress_type=item.compress_type)

    print(f"[OK] Gerado: {output_pbix}")

# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    SOURCE = r"CAMINHO\DO\TEMPLATE.pbix"   # template com DataModel local
    OUTPUT = r"CAMINHO\DE\SAIDA.pbix"

    pages = [page_exemplo()]
    layout = build_layout(pages)
    pack_pbix(SOURCE, OUTPUT, layout)
    print("Concluido!")
```

---

## Inspecionar campos de um Excel

Script auxiliar para descobrir nomes exatos de campos antes de montar o painel:

```python
import zipfile, xml.etree.ElementTree as ET

xlsx_path = r"CAMINHO\PARA\PLANILHA.xlsx"

with zipfile.ZipFile(xlsx_path, 'r') as z:
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    ss = ET.fromstring(z.read('xl/sharedStrings.xml'))
    ss_list = [''.join(t.text or '' for t in si.findall('.//m:t', ns))
               for si in ss.findall('m:si', ns)]

    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = [(s.get('name'), s.get('sheetId'))
              for s in wb.findall('.//m:sheet', ns)]

    for sheet_name, sheet_id in sheets:
        sheet_file = f'xl/worksheets/sheet{sheet_id}.xml'
        try:
            sheet_xml = ET.fromstring(z.read(sheet_file))
        except KeyError:
            continue
        rows = sheet_xml.findall('.//m:row', ns)
        if not rows:
            continue
        headers = []
        for c in rows[0].findall('m:c', ns):
            v = c.find('m:v', ns)
            if v is not None:
                val = ss_list[int(v.text)] if c.get('t') == 's' else v.text
                headers.append(repr(val))
        print(f"{repr(sheet_name)}: {headers}")
```

---

## Cores padrão recomendadas

| Uso | Hex |
|-----|-----|
| Header principal | `#15314F` |
| Faixa de slicer | `#1B3F60` |
| Fundo da página | `#EEF1F6` |
| Fundo do card | `#F2F2F2` |
| Texto principal | `#15314F` |
| Label secundário | `#555555` |
| Texto sobre azul | `#FFFFFF` |
| Detalhe azul claro | `#A9C0D8` |

---

## Dimensões de referência (canvas 1280×720)

| Elemento | X | Y | W | H |
|---------|---|---|---|---|
| Header | 0 | 0 | 1280 | 66 |
| Faixa slicers | 0 | 66 | 1280 | 42 |
| 4 cards (row) | 10/315/620/925 | 116 | 295/295/295/345 | 90 |
| Gráfico grande | 10 | 216 | 1260 | 300 |
| 2 gráficos lado a lado | 10 / 640 | 216 | 620 / 630 | 296 |
| Barra de rodapé | 10 | 522 | 1260 | 186 |
