# pptx-mcp-server

Servidor MCP (stdio) que converte um arquivo HTML — escrito inteiramente pela LLM chamadora —
em um PowerPoint (`.pptx`) **nativo e editável**. O servidor nunca gera nem embute nenhuma
imagem/screenshot do slide inteiro; ele só usa um Chromium headless para *medir* o layout real
(posições, fontes, cores) do HTML e converte o que foi marcado em objetos de verdade do PPTX
(caixas de texto, formas, imagens).

## Como funciona

1. **A LLM escreve tudo.** Todo o design, conteúdo e HTML/CSS (e imagens locais, se houver) são
   escritos pela LLM chamadora diretamente no workspace do usuário — o servidor não desenha, não
   escolhe layout, não tem opinião de design. Ele só converte.
2. **Contrato de marcação.** Cada slide é um elemento HTML com `data-pptx-slide`; dentro dele,
   qualquer elemento com `data-pptx="text"`, `data-pptx="shape"` ou `data-pptx="image"` vira um
   objeto real no `.pptx`. Tudo que não tem `data-pptx` é só andaime de layout (divs de flexbox,
   grid, etc.) e é ignorado na conversão. Veja a tool `get_pptx_authoring_guide` para o contrato
   completo, com exemplo.
3. **Conversão.** O servidor abre o HTML num Chromium headless (só para ler
   `getBoundingClientRect()`/`getComputedStyle()` de cada elemento marcado — nenhum screenshot é
   tirado) e usa esses dados para montar o `.pptx` via `pptxgenjs`: texto vira caixa de texto
   editável, formas viram retângulos/retângulos arredondados com fill/borda reais, imagens viram
   objetos de imagem nativos.

## Tools expostas

- `get_pptx_authoring_guide()` — devolve o guia completo de como escrever o HTML (o contrato
  `data-pptx-slide`/`data-pptx`, o que cada tipo lê de CSS, o que não é suportado, e orientações
  de design não-restritivas). **Chame antes de escrever qualquer HTML** — a conversão depende
  desse contrato; HTML sem essas marcações vira um `.pptx` vazio.
- `convert_html_to_pptx({ htmlPath, outputPath? })` — **a tool que entrega o arquivo**. Recebe o
  caminho absoluto de um único arquivo HTML já escrito pela LLM, extrai os elementos marcados e
  grava o `.pptx`. Devolve o caminho salvo e, se houver, avisos sobre elementos que não
  converteram bem (tamanho zero, imagem não encontrada, slide com tamanho diferente do primeiro).

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
