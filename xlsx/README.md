# Pasta para Arquivos Excel

Esta pasta é onde você deve colocar seus arquivos Excel (.xlsx, .xlsm) que deseja ler através do MCP server.

## Segurança

Por motivos de segurança, o MCP server **só pode acessar arquivos dentro desta pasta** `xlsx/` e suas subpastas.

## Estrutura Recomendada

Você pode organizar seus arquivos em subpastas:

```
xlsx/
├── exemplo.xlsx
├── vendas/
│   ├── 2024.xlsx
│   └── 2025.xlsx
├── clientes/
│   └── dados_clientes.xlsx
└── relatorios/
    └── mensal.xlsx
```

## Como Referenciar

Ao usar as ferramentas do MCP, use caminhos relativos a esta pasta:

- `"exemplo.xlsx"` → acessa `xlsx/exemplo.xlsx`
- `"vendas/2024.xlsx"` → acessa `xlsx/vendas/2024.xlsx`
- `"relatorios/mensal.xlsx"` → acessa `xlsx/relatorios/mensal.xlsx`

## Criar Arquivo de Exemplo

Para testar o MCP server, coloque um arquivo Excel de exemplo nesta pasta e use o VS Code com GitHub Copilot.
