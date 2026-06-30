# powerbi-generator

Plugin para geração de painéis Power BI (`.pbix`) via Python, sem abrir o Power BI Desktop.

## Como funciona

Um `.pbix` é um arquivo ZIP com `Report/Layout` (JSON UTF-16 LE) e `DataModel` (binário SSAS). Este plugin encapsula o processo validado de substituir o Layout e contornar o `SecurityBindings` (blob DPAPI), que foi descoberto por trial-and-error no projeto SST do SENAC MT.

## Skills disponíveis

### `/gerar-pbix <projeto>` — comando do usuário

Guia interativo para gerar um painel novo. Claude vai:
1. Coletar template, fonte de dados, páginas e visuais desejados
2. Inspecionar os campos exatos do Excel
3. Gerar o script Python adaptado ao projeto
4. Executar e entregar o `.pbix` pronto

### `pbix-context` — skill automática

Ativada quando o assunto é geração de `.pbix` via Python. Injeta as regras críticas no contexto para evitar os erros já conhecidos.

## Requisitos

- Python 3.x (biblioteca padrão apenas — sem dependências externas)
- Um arquivo `.pbix` template com DataModel local (sem `RemoteArtifacts`)
- Power BI Desktop (para abrir o resultado)

## Referência técnica

Ver [`skills/gerar-pbix/references/template-script.md`](skills/gerar-pbix/references/template-script.md) para o script completo com todas as funções de visual.
