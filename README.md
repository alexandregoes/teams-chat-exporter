# Teams Chat Exporter (Extensão Chrome MV3)

Extensão para Google Chrome que exporta o histórico de um chat específico aberto no **Microsoft Teams Web** (`teams.microsoft.com`, `teams.live.com` ou `teams.cloud.microsoft`) para **HTML (`.html`)**, **Markdown (`.md`)**, **JSON (`.json`)** ou **Texto simples (`.txt`)**.

## Como instalar no Chrome

1. Abra `chrome://extensions/` no Google Chrome.
2. Ative o **Modo do desenvolvedor** (canto superior direito).
3. Clique em **Carregar sem compactação** (*Load unpacked*).
4. Selecione a pasta `teams-chat-exporter`.

## Como usar

1. Abra o Microsoft Teams no navegador e entre no **chat específico** que deseja exportar.
2. Clique no ícone da extensão **Teams Chat Exporter** na barra do Chrome.
3. Configure as opções desejadas:
   - **Rolar para cima automaticamente**: carrega mensagens antigas enquanto coleta (recomendado para chats longos).
   - **Limite de mensagens**: deixe `0` para exportar tudo ou defina um número máximo.
   - **Formato**: escolha `.html`, `.md`, `.json` ou `.txt`.
4. Clique em **Coletar Chat** e aguarde a conclusão (ou clique em **Parar** a qualquer momento).
5. Clique em **Baixar Arquivo**.
