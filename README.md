# Jspreadsheet OpenAI extension

`@jspreadsheet/openai` adds the spreadsheet formula **`=PROMPT(...)`** to Jspreadsheet.
The formula sends its arguments to the OpenAI chat completions API and writes the
answer into the cell, so a spreadsheet can ask things like:

```
=PROMPT("Capital of", A1)
=PROMPT("Translate to French:", B2)
=PROMPT("Classify this support ticket as billing, bug or other:", C3)
```

The prompt is **evaluated by the Jspreadsheet server**, never by the browser. The API key
stays on the server, every collaborator sees the same answer, and the answer is cached
with the document so reloading a spreadsheet does not call the API again.

## How it works

```
 browser A ──┐                       ┌─ OpenAI API
 browser B ──┼── socket.io ── @jspreadsheet/server ──┤
 browser C ──┘                       └─ document storage (cache)
```

1. The extension is registered on **both** sides with `jspreadsheet.setExtensions({ openai })`
   so `PROMPT` is a known formula everywhere. An unknown function would render `#NAME?`.
2. In the **browser** the formula only returns the `...` placeholder. The OpenAI SDK is not
   loaded in the browser at all.
3. On the **server** the document is loaded headless with the same engine. There `PROMPT`
   calls the API, stores the answer in the cell cache with `worksheet.setCache()` and
   broadcasts that call to every browser that has the document open. Each peer replays it
   and the `...` turns into the answer.
4. The cache is persisted with the document. When the document is loaded again, the cached
   answer is reused and the API is not billed a second time. Editing the formula or one of
   the cells it references recalculates it and asks the API again.

## Requirements

- A Jspreadsheet Pro license (client id and license key). Get a free trial at
  https://jspreadsheet.com/me. The same key is used by the server and the browser.
- Jspreadsheet 13 and the `@jspreadsheet/formula-pro` extension (the formula engine).
- A running `@jspreadsheet/server` (see the example below or the demo repository).
- An OpenAI API key: https://platform.openai.com/api-keys

## Installation

On the server:

```bash
npm install @jspreadsheet/openai
```

The `openai` SDK is a dependency of the package and is installed with it.

In the browser the same package is used, but only to register the formula name.
If you build with webpack or Vite, map the `openai` module to nothing so the SDK is not
bundled:

```js
// webpack.config.js
resolve: { fallback: { openai: false } }
```

If you use the `intrasheets` package, the extension is already included and registered.
Nothing to install on the browser side.

## Server example

A minimal collaboration server with MongoDB persistence, the REST API and the OpenAI
extension. The server holds the API key and evaluates every `PROMPT`.

```bash
npm install @jspreadsheet/server @jspreadsheet/server-api @jspreadsheet/server-mongodb @jspreadsheet/openai dotenv
```

```js
// index.js
require('dotenv').config();

const server = require('@jspreadsheet/server');
const api = require('@jspreadsheet/server-api');
const adapter = require('@jspreadsheet/server-mongodb');
const openai = require('@jspreadsheet/openai');

// 1. Create the OpenAI client. Call this once per process.
openai({
    apiKey: process.env.OPENAI_API_KEY,
    // Optional: change the request before it is sent (model, temperature, messages...)
    onbeforesend: function(instance, request) {
        request.model = 'gpt-4o-mini';
    },
    // Optional: post-process the answer before it lands in the cell
    onsuccess: function(value, result, x, y, worksheet) {
        return value.trim();
    },
});

// 2. Start the server with the extension registered.
//    The server passes its core to the extension so it can broadcast the answers.
server({
    port: 3000,
    config: {
        cors: { origin: '*' },
    },
    // Persistence: delegate the document hooks to an adapter (mongodb, pgsql, redis).
    // The cache written by PROMPT is saved with the document through these hooks.
    load: async function(guid, auth) {
        return await adapter.load(guid, auth);
    },
    change: async function(guid, changes, auth, onerror) {
        return await adapter.change(guid, changes, auth, onerror);
    },
    create: async function(guid, config, auth) {
        return await adapter.create(guid, config, auth);
    },
    destroy: async function(guid, auth) {
        return await adapter.destroy(guid, auth);
    },
    replace: async function(guid, config, auth) {
        return await adapter.replace(guid, config, auth);
    },
    list: async function(auth) {
        return await adapter.list(auth);
    },
    extensions: { api, openai },
    license: {
        clientId: process.env.JSS_CLIENT,
        licenseKey: process.env.JSS_LICENSE,
    },
    error: function(e) {
        console.error(e);
    },
});
```

```ini
# .env
JSS_CLIENT=your-client-id
JSS_LICENSE=your-license-key
OPENAI_API_KEY=sk-...
```

The example has no authentication: add the `beforeConnect`, `beforeLoad` and `beforeChange`
hooks before exposing it. See https://jspreadsheet.com/docs/server for the full server
documentation, including authentication, adapters and scaling. The demo repository below
is a complete, authenticated version of this file.

## Browser example

The browser registers the same extensions and connects to the server with the client
extension. The `PROMPT` formula returns `...` until the server answer arrives.

```html
<script src="https://cdn.socket.io/4.8.0/socket.io.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jsuites@6/dist/jsuites.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspreadsheet@13/dist/index.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@jspreadsheet/formula-pro/dist/index.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@jspreadsheet/client/dist/index.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@jspreadsheet/openai/dist/index.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/jsuites@6/dist/jsuites.css" />
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/jspreadsheet@13/dist/jspreadsheet.css" />

<div id="spreadsheet"></div>

<script>
jspreadsheet.setLicense({ clientId: 'your-client-id', licenseKey: 'your-license-key' });

// Same extensions as the server: formula engine, realtime client and PROMPT
jspreadsheet.setExtensions({ formula, client, openai });

// Connect to the server
client({
    url: 'http://localhost:3000',
    token: 'jwt-of-the-signed-in-user',
});

// Open a document that exists on the server
jspreadsheet(document.getElementById('spreadsheet'), {
    guid: 'document-guid',
});
</script>
```

Or with modules:

```js
import jspreadsheet from 'jspreadsheet';
import formula from '@jspreadsheet/formula-pro';
import client from '@jspreadsheet/client';
import openai from '@jspreadsheet/openai';

jspreadsheet.setExtensions({ formula, client, openai });
```

Type `=PROMPT("Capital of", "France")` into a cell. The cell shows `...`, the server
evaluates the prompt, and a moment later every open browser shows `Paris`.

## Options

`openai(options)` is called once on the server. The hooks below are removed from the
object; every other property is passed untouched to the OpenAI SDK constructor
(`new OpenAI(options)`), so `baseURL`, `organization`, `fetch`, `timeout` and the other
SDK options all work.

| Option         | Type       | Description |
|----------------|------------|-------------|
| `apiKey`       | `string`   | **Required.** OpenAI API key. Without it the extension stays disconnected and `PROMPT` logs `Not connected...`. |
| `onbeforesend` | `function` | `(instance, request)` called before each request. Mutate `request` in place to change the model, add system messages, set `temperature`, etc. |
| `onsuccess`    | `function` | `(value, result, x, y, worksheet)` called with the text answer and the raw API result. Return `false` to drop the answer, return a value to replace it, or return nothing to keep it. May be `async`. |

The default request is one system message containing all the formula arguments joined
with a space, sent to the `gpt-4` model:

```js
{
    model: 'gpt-4',
    messages: [{ role: 'system', content: 'Capital of France' }]
}
```

Use `onbeforesend` to pick a different model or to prepend instructions:

```js
onbeforesend: function(instance, request) {
    request.model = 'gpt-4o-mini';
    request.messages.unshift({
        role: 'system',
        content: 'Answer with a single value, no explanation.'
    });
}
```

## Formulas

| Formula                 | Description |
|-------------------------|-------------|
| `PROMPT(text1, text2, ...)` | Joins all arguments with a space and sends them as the prompt. Arguments can be strings or cell references, which the engine resolves before the call. Returns `...` until the server answer arrives. |
| `SUMMARY(...)`          | Reserved, not implemented yet. |

## Demo application

A complete, dockerised reference application (React + `@jspreadsheet/server` + MongoDB)
lives at **https://github.com/intrasheets/demo**. It runs the `intrasheets` package, which
already includes this extension on the browser side.

```bash
git clone https://github.com/intrasheets/demo.git intrasheets-demo
cd intrasheets-demo
cp .env.example .env     # set JSS_CLIENT, JSS_LICENSE, JWT_SECRET
docker compose up
```

To enable `PROMPT` in the demo, add `OPENAI_API_KEY` to `.env`, install
`@jspreadsheet/openai` in `server/`, and register it in `server/index.js`:

```js
const openai = require('@jspreadsheet/openai');

openai({ apiKey: process.env.OPENAI_API_KEY });

server({
    // ...
    extensions: { api, application, openai },
});
```

## Troubleshooting

- **`#NAME?` in the cell**: the extension is not registered in the browser, or the formula
  engine (`@jspreadsheet/formula-pro`) is missing. `setExtensions` must include `formula`
  before or together with `openai`.
- **The cell stays `...`**: the server did not answer. Check that `openai({ apiKey })` was
  called on the server, that the server was started with `extensions: { openai }`, and
  look at the server log for `Not connected...` or an API error.
- **The API is called on every reload**: the cache is stored with the document, so the
  document must be persisted by a server adapter. Documents opened without a server, or
  without persistence, have no cache to reuse.
- **I want a different model**: set `request.model` in `onbeforesend`.

## Links

- Jspreadsheet: https://jspreadsheet.com
- Server documentation: https://jspreadsheet.com/docs/server
- Server extensions: https://jspreadsheet.com/docs/server/extensions
- Demo application: https://github.com/intrasheets/demo
- Issues: https://github.com/jspreadsheet/pro/issues

## License

This extension requires a valid Jspreadsheet Pro license. See https://jspreadsheet.com/pricing.
