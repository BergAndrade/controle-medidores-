# Controle de Medidores

Aplicação web para controle acumulativo de instalação e retirada de medidores.

## Recursos

- Importação de bases Excel (`.xlsx`, `.xls` e `.csv`)
- Banco local acumulativo com IndexedDB
- Dashboard com filtros por mês, data, fase e tipo de nota
- Consulta de medidores em massa
- Validação de medidores instalados com menos de 11 dígitos
- Histórico de importações com exclusão por lote
- Exportação do consolidado para Excel
- Backup e restauração em JSON

## Estrutura

```text
controle-medidores/
├── index.html
├── css/
│   └── style.css
├── js/
│   └── app.js
├── assets/
├── README.md
└── .gitignore
```

## Executar localmente

Abra `index.html` em um navegador moderno. Para desenvolvimento, também é possível servir a pasta com um servidor HTTP local.

## Publicar no GitHub Pages

1. Crie um repositório no GitHub.
2. Envie todo o conteúdo desta pasta para a branch `main`.
3. No repositório, abra **Settings > Pages**.
4. Em **Build and deployment**, escolha **Deploy from a branch**.
5. Selecione `main` e `/ (root)` e salve.

> Os registros operacionais ficam no IndexedDB do navegador/dispositivo. Publicar o app não envia automaticamente esses dados ao GitHub. Faça backups periódicos pelo próprio app.

## Bibliotecas

A aplicação usa SheetJS/XLSX e Chart.js via CDN.
