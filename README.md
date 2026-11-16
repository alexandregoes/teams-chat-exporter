# Teams Chat Exporter

Extensão para Google Chrome (Manifest V3) que exporta o histórico completo de um chat específico do **Microsoft Teams Web** (`teams.microsoft.com`, `teams.live.com` e `teams.cloud.microsoft`) em **HTML** com visual fiel ao Teams, além de **Markdown (`.md`)**, **JSON (`.json`)** e **Texto simples (`.txt`)**.

## Recursos

- **Visual fiel ao Microsoft Teams (HTML):**
  - Tema escuro padrão do Teams, com botão no cabeçalho para alternar entre **Tema claro** e **Tema escuro**.
  - Ordem cronológica correta (**mensagens mais antigas no topo**, mais recentes embaixo), com divisores de data e agrupamento de mensagens consecutivas.
  - Barra de busca integrada no HTML exportado (`Buscar no chat...`).
- **Coleta completa de conteúdo:**
  - **Fotos de perfil (avatares)** e **imagens das mensagens** embutidas em Base64 (funciona offline).
  - **Emoticons** Unicode e animados do Teams, além de **reações** (ex.: `👍 1`).
  - **Tabelas** preservadas com rolagem horizontal para colunas largas.
  - **Respostas citadas (replies)** em cartão destacado, **`@menções`** inline, links clicáveis, formatação (`negrito`, `itálico`, `código`, listas), badge `URGENT` e cartões de **anexos/arquivos**.
- **Rolagem automática inteligente:** vai primeiro até o final do chat para garantir as mensagens mais novas e depois sobe todo o histórico automaticamente.

## Instalação

1. Baixe ou clone este repositório:
   ```bash
   git clone https://github.com/alexandregoes/teams-chat-exporter.git
   ```
2. Abra `chrome://extensions/` no Google Chrome.
3. Ative o **Modo do desenvolvedor** (canto superior direito).
4. Clique em **Carregar sem compactação** (*Load unpacked*) e selecione a pasta `teams-chat-exporter`.

## Como usar

1. Abra o Microsoft Teams no navegador e entre no **chat específico** que deseja exportar.
2. Clique no ícone da extensão **Teams Chat Exporter** na barra do Chrome.
3. Configure as opções:
   - **Rolar para cima automaticamente**: carrega todo o histórico desde o início do chat.
   - **Incluir imagens (embutidas no HTML)**: salva imagens, emoticons e fotos de perfil em Base64 dentro do HTML.
   - **Limite de mensagens**: deixe `0` para exportar todas ou informe um limite máximo.
   - **Formato**: escolha `.html`, `.md`, `.json` ou `.txt`.
4. Clique em **Coletar Chat** e aguarde a conclusão (ou clique em **Parar** a qualquer momento).
5. Clique em **Baixar Arquivo**.

## Estrutura do projeto

- `manifest.json` — configuração Manifest V3 da extensão.
- `content.js` — script injetado no Teams Web responsável pela rolagem, extração e sanitização das mensagens.
- `popup.html` / `popup.css` / `popup.js` — interface do popup e gerador dos arquivos HTML, Markdown, JSON e TXT.
- `LICENSE` — licença do projeto.

## Licença

Este projeto está licenciado sob a [Creative Commons BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/) (**Atribuição - Não Comercial - Compartilha Igual**). Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
