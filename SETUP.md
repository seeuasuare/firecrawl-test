# Установка MCP Firecrawl для Claude Code

## Проблема

При попытке выполнить `claude mcp add firecrawl ...` терминал выдаёт ошибку
"command not found", потому что Claude Code CLI не установлен.

## Решение

### Шаг 1. Установить Node.js v20+

Проверьте версию:

```bash
node -v
```

Если не установлен — скачайте с https://nodejs.org/

### Шаг 2. Установить Claude Code CLI

```bash
npm install -g @anthropic-ai/claude-code
```

Проверка:

```bash
claude --version
```

> Если ошибка прав доступа на Linux/macOS:
>
> ```bash
> sudo npm install -g @anthropic-ai/claude-code
> ```
>
> Или настройте npm prefix:
>
> ```bash
> mkdir -p ~/.npm-global
> npm config set prefix '~/.npm-global'
> echo 'export PATH=~/.npm-global/bin:$PATH' >> ~/.bashrc
> source ~/.bashrc
> npm install -g @anthropic-ai/claude-code
> ```

### Шаг 3. Получить API-ключ Firecrawl

Зарегистрируйтесь и получите ключ: https://www.firecrawl.dev/app/api-keys

### Шаг 4. Добавить Firecrawl MCP-сервер

```bash
claude mcp add firecrawl -s user -e FIRECRAWL_API_KEY=fc-YOUR_API_KEY -- npx -y firecrawl-mcp
```

### Шаг 5. Проверить

```bash
claude mcp list
```

## Альтернатива: ручная настройка без CLI

Создайте или отредактируйте файл `~/.claude.json`:

```json
{
  "mcpServers": {
    "firecrawl": {
      "command": "npx",
      "args": ["-y", "firecrawl-mcp"],
      "env": {
        "FIRECRAWL_API_KEY": "fc-YOUR_API_KEY"
      }
    }
  }
}
```

## Ссылки

- [Документация Firecrawl MCP Server](https://docs.firecrawl.dev/mcp-server)
- [Firecrawl MCP Server на GitHub](https://github.com/firecrawl/firecrawl-mcp-server)
- [Claude Code](https://docs.anthropic.com/en/docs/claude-code)
