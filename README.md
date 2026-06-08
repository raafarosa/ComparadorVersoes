# Comparador de Versões

Um aplicativo web leve para comparar versões de aplicações entre ambientes de homologação (HML), pré-produção (PREPROD) e produção (PROD).

## O que faz

- Detecta automaticamente os ambientes a partir dos nomes dos arquivos ou permite seleção manual por ambiente.
- Processa arquivos JSON, planilhas Excel (`.xlsx`, `.xls`) e texto tabulado comum.
- Normaliza nomes de aplicações e ignora registros de backup, pastas antigas e versões duplicadas.
- Compara as versões entre os ambientes e apresenta o resultado em tabelas interativas.
- Exporta o resultado da comparação em CSV.

## Como usar

1. Abra `index.html` em um navegador compatível.
2. Escolha o modo:
   - **Detecção Automática**: selecione arquivos de qualquer lugar e o app tentará identificar HML, PREPROD e PROD pelo nome.
   - **Modo Manual**: faça upload separado para cada ambiente.
3. Clique em **Comparar Ambientes**.
4. Navegue entre as abas:
   - PREPROD x PROD
   - PREPROD x HML
   - HML x PROD
   - Duplicadas Detectadas
5. Opcional: clique em **Exportar CSV** para salvar o relatório.

## Formatos aceitos

- Arquivos JSON com estrutura similar a:
  - `Aplicacoes` / `aplicacoes` / `Aplicações`
  - `Path`, `Changeset`, `Release`
- Arquivos Excel (`.xlsx`, `.xls`) contendo colunas de aplicação, changeset e release.
- Arquivos de texto tabulado com colunas separadas por tabulação ou espaços.

## Arquivos do projeto

- `index.html` - interface web do comparador.
- `styles.css` - estilos visuais do app.
- `script.js` - lógica de processamento, leitura de arquivos e comparação.

## Observações

- O app é executado no cliente, sem necessidade de backend.
- A detecção de versões trata casos como `C123 | R456`, `C123` e `Sem versão`.
- Registros duplicados são identificados e agrupados na aba correspondente.

## Requisitos

- Navegador moderno com suporte a JavaScript e APIs de arquivos.

## Desenvolvimento

Basta editar os arquivos `index.html`, `styles.css` e `script.js` e recarregar a página.

