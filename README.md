# pptx-mcp-server

Servidor MCP (stdio) que renderiza slides em HTML/CSS e exporta para um arquivo PowerPoint (`.pptx`).

## Como funciona

1. Cada slide é um documento HTML completo e independente (com `<style>` inline), dimensionado
   para o tamanho do slide (padrão 1280x720 px, 16:9).
2. O servidor renderiza cada HTML com Chromium headless (Playwright) e captura um screenshot PNG.
3. As imagens são montadas em um `.pptx` (via `pptxgenjs`), uma por slide, ocupando o slide inteiro.

Como o slide final é uma imagem, o texto **não é editável** dentro do PowerPoint — a prioridade
dessa abordagem é fidelidade visual ao design HTML/CSS, não edição posterior no PowerPoint.

## Tools expostas

- `preview_slide({ html, width?, height? })` — **só para checagem visual**, não salva arquivo.
  Renderiza um único slide (1x escala) e devolve a imagem PNG em base64, para iterar no design
  antes de gerar o arquivo final.
- `build_pptx({ slides: string[], outputPath?, width?, height? })` — **a tool que entrega o
  arquivo**. Renderiza todos os slides (2x escala, mais nítido), em ordem, e grava o `.pptx` em
  disco.

### Resolução de `outputPath`

1. Caminho absoluto informado → usado diretamente.
2. Caminho relativo ou omitido → tenta a primeira `root` declarada pelo cliente MCP (se ele
   suportar o recurso `roots` do protocolo).
3. Se o cliente não suportar `roots` → usa a variável de ambiente `PPT_MCP_OUTPUT_DIR`, se
   definida, ou `~/Documents/PPT-MCP` como último fallback.

## Uso

```bash
npm install   # também baixa o Chromium do Playwright (postinstall)
npm run build
npm start     # inicia o servidor MCP via stdio
```

## Configuração nos principais harnesses/clientes MCP

O pacote está publicado no npm como [`pptx-mcp-server`](https://www.npmjs.com/package/pptx-mcp-server),
então na maioria dos clientes basta apontar `command: npx`, `args: ["-y", "pptx-mcp-server"]`. A
env var `PPT_MCP_OUTPUT_DIR` é opcional (ver [Resolução de `outputPath`](#resolução-de-outputpath));
troque `C:\caminho\padrao\de\saida` pela pasta que preferir.

### Claude Code

```bash
claude mcp add pptx-mcp-server -e PPT_MCP_OUTPUT_DIR=C:\caminho\padrao\de\saida -- npx -y pptx-mcp-server
```

Ou editando `.mcp.json` (projeto) / config de usuário diretamente — mesmo formato JSON do bloco
"Claude Desktop / Cursor" abaixo.

### Claude Desktop

Edite `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "pptx-mcp-server": {
      "command": "npx",
      "args": ["-y", "pptx-mcp-server"],
      "env": {
        "PPT_MCP_OUTPUT_DIR": "C:\\caminho\\padrao\\de\\saida"
      }
    }
  }
}
```

### Cursor

Mesmo formato do Claude Desktop, em `~/.cursor/mcp.json` (global) ou `.cursor/mcp.json` (projeto):

```json
{
  "mcpServers": {
    "pptx-mcp-server": {
      "command": "npx",
      "args": ["-y", "pptx-mcp-server"],
      "env": {
        "PPT_MCP_OUTPUT_DIR": "C:\\caminho\\padrao\\de\\saida"
      }
    }
  }
}
```

### OpenAI Codex CLI

Adicione em `~/.codex/config.toml` (ou `.codex/config.toml` do projeto):

```toml
[mcp_servers.pptx-mcp-server]
command = "npx"
args = ["-y", "pptx-mcp-server"]
env = { PPT_MCP_OUTPUT_DIR = "C:\\caminho\\padrao\\de\\saida" }
```

### Google Antigravity (CLI `agy` / IDE)

Antigravity 2.0, a IDE e a CLI compartilham a mesma config, em
`~/.gemini/config/mcp_config.json` (ou `.agents/mcp_config.json` no workspace). Mesmo formato
`mcpServers`/`command`/`args` de cima. Reinicie o Antigravity depois de editar.

### VS Code (GitHub Copilot Chat, modo agente)

Crie/edite `.vscode/mcp.json` no projeto — repare que o VS Code usa a chave `servers` (não
`mcpServers`) e exige `"type": "stdio"`:

```json
{
  "servers": {
    "pptx-mcp-server": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "pptx-mcp-server"],
      "env": {
        "PPT_MCP_OUTPUT_DIR": "C:\\caminho\\padrao\\de\\saida"
      }
    }
  }
}
```

### GitHub Copilot CLI

Via wizard interativo (`/mcp add` dentro do `copilot`) ou editando `~/.copilot/mcp-config.json`:

```json
{
  "mcpServers": {
    "pptx-mcp-server": {
      "type": "local",
      "command": "npx",
      "args": ["-y", "pptx-mcp-server"],
      "tools": ["*"],
      "env": {
        "PPT_MCP_OUTPUT_DIR": "C:\\caminho\\padrao\\de\\saida"
      }
    }
  }
}
```

### Windsurf

Mesmo formato `mcpServers`/`command`/`args`, em `~/.codeium/windsurf/mcp_config.json`.

### Zed

Em `settings.json` (`~/.config/zed/settings.json` no macOS/Linux, `%APPDATA%\Zed\settings.json`
no Windows), a chave é `context_servers` e servidores customizados precisam de `"source": "custom"`:

```json
{
  "context_servers": {
    "pptx-mcp-server": {
      "source": "custom",
      "command": "npx",
      "args": ["-y", "pptx-mcp-server"],
      "env": {
        "PPT_MCP_OUTPUT_DIR": "C:\\caminho\\padrao\\de\\saida"
      }
    }
  }
}
```

### Outros clientes

Quase todo cliente MCP stdio segue o mesmo formato básico — uma entrada com `command: "npx"` e
`args: ["-y", "pptx-mcp-server"]`, variando só a chave de agrupamento (`mcpServers`, `servers`,
`context_servers`) e o caminho do arquivo de config. Consulte a documentação do cliente específico
se ele não estiver nesta lista.

### Rodando localmente (sem publicar/instalar via npm)

Durante desenvolvimento, ou se preferir não depender do registro npm, aponte `command` direto pro
build local em vez de `npx`:

```json
{
  "mcpServers": {
    "pptx-mcp-server": {
      "command": "node",
      "args": ["C:\\Projetos\\mcps\\ppt\\dist\\server.js"],
      "env": {
        "PPT_MCP_OUTPUT_DIR": "C:\\caminho\\padrao\\de\\saida"
      }
    }
  }
}
```
