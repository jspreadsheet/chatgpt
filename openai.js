/**
 * Jspreadsheet OpenAI extension
 *
 * Adds the spreadsheet formula =PROMPT("text", ...) which sends its
 * arguments to the OpenAI chat completions API and writes the answer into
 * the cell.
 *
 * How it works (the prompt is evaluated by the SERVER, never by the browser):
 *
 *   1. The same file is registered as an extension on both sides with
 *      jspreadsheet.setExtensions({ openai }) so that PROMPT is a known
 *      formula everywhere (an unknown function would render #NAME?).
 *   2. In the browser the formula only returns the '...' placeholder. The
 *      OpenAI SDK is not even loaded there (the bundlers map the 'openai'
 *      module to nothing), so the API key never reaches the client.
 *   3. On the server (@jspreadsheet/server) the document is loaded headless
 *      with the same engine. There, PROMPT calls the API, stores the answer
 *      in the cell cache with worksheet.setCache(), and broadcasts a
 *      { method: 'setCache' } message to every socket in the document's
 *      guid room. The client extension replays that call on each peer, so
 *      every open browser sees the answer in the cell.
 *   4. The cache is persisted with the document. When a document is loaded
 *      again (spreadsheet.building === true) the cached answer is reused and
 *      the API is not called a second time.
 *
 * Usage on the server:
 *
 *   const openai = require('@jspreadsheet/openai');
 *   // Create the API client (the extension is a singleton: one client per process)
 *   openai({
 *       apiKey: process.env.OPENAI_API_KEY,         // forwarded to new OpenAI(options)
 *       onbeforesend: function(instance, request) {}, // mutate the request (model, messages...)
 *       onsuccess: function(value, result, x, y, worksheet) {}, // post-process the answer
 *   });
 *   // Register the formula on the engine; the server passes its core as 2nd argument
 *   jspreadsheet.setExtensions({ openai }, serverCore);
 *
 * Any other property of the options object is passed straight to the OpenAI
 * SDK constructor (baseURL, organization, fetch, ...).
 *
 * Tests: server/tests/openai.test.js (runs the real engine headless with a
 * fake transport and a fake server core).
 */

// Node only: load the OpenAI SDK unless a global OpenAI client was already
// provided by the host. `var OpenAI` is hoisted, so the check does not throw
// in the browser, where require() is not a function and OpenAI stays undefined.
if (! OpenAI && typeof(require) === 'function') {
    var OpenAI = require('openai');
}

// UMD wrapper: CommonJS (server), AMD, or a global `openai` in the browser
;(function(global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory() :
    typeof define === 'function' && define.amd ? define(factory) :
    global.openai = factory();
}(this, (function() {

    // Server core handed to Extension.license() by @jspreadsheet/server.
    // Only two things are used from it: helpers.encodeZip() to compress the
    // broadcast payload and io.to(room).emit() to reach the sockets of a document.
    let Server = null;
    // The jspreadsheet namespace the extension was registered on (captured in license())
    let JSS = null;
    // OpenAI SDK client, created by Extension(options). Null until configured,
    // and always null in the browser.
    let instance = null;
    // Optional hook onbeforesend(instance, request): called with the request
    // object before it is sent, so the host can change the model, add
    // messages, temperature, etc. Mutate `request` in place.
    let beforeSend;
    // Optional hook onsuccess(value, result, x, y, worksheet): called with the
    // text answer and the raw API result. Return false to drop the answer,
    // return a value to replace it, or return undefined to keep it.
    let success;

    /**
     * Send a prompt to the chat completions API and hand the answer to `done`.
     *
     * Must be called with `this` = { w: worksheet, x, y } so the hooks know
     * which cell is being evaluated.
     *
     * @param {string[]} prompt - formula arguments, joined with a space into one message
     * @param {function} done - async callback receiving the final text value
     */
    const query = function(prompt, done) {
        if (instance) {
            // Default request: one system message with the whole prompt.
            // onbeforesend can rewrite anything here, including the model.
            let request = {
                messages: [{
                    role: "system",
                    content: prompt.join(' ')
                }],
                model: "gpt-4",
            };

            if (beforeSend) {
                beforeSend(instance, request);
            }

            // Cell being evaluated (worksheet instance and coordinates)
            let w = this.w;
            let x = this.x;
            let y = this.y;

            // Call API. The promise is intentionally not returned: the
            // formula already answered with the placeholder and the result
            // arrives later through the cache broadcast.
            instance.chat.completions.create(request).then(async function(result) {
                // Plain text of the first choice
                let value = result.choices[0].message.content;
                if (success) {
                    let ret = await success(value, result, x, y, w);
                    if (ret === false) {
                        // Host vetoed the answer: leave the cell as it is
                        return;
                    } else {
                        if (typeof(ret) !== 'undefined') {
                            // Host replaced the answer
                            value = ret;
                        }
                    }
                }
                await done(value);
            });
        } else {
            // Extension(options) was never called with an apiKey
            console.log('Not connected...');
        }
    }

    /**
     * True in the browser, false on the server (headless engine has no document)
     */
    const isFront = function() {
        return typeof(document) === 'object';
    }

    /**
     * Build the formula map registered on the engine through
     * JSS.formula.setFormula(). Each entry becomes a spreadsheet function.
     *
     * Formulas run with `this` = { instance: worksheet, x, y } provided by
     * the formula engine.
     */
    const getFormulas = function() {
        const OPENAI = {}

        /**
         * =PROMPT(text1, text2, ...)
         *
         * All arguments are concatenated with spaces into one prompt. Cell
         * references are resolved by the engine before this is called, so
         * =PROMPT("Summarize:", A1) works.
         *
         * Returns '...' immediately. The real answer is written to the cell
         * cache asynchronously by the server, see the file header.
         */
        OPENAI.PROMPT = function() {
            // Worksheet instance and coordinates of the cell being calculated
            let w = this.instance;
            let x = this.x;
            let y = this.y;
            // Cell record
            let cell = w.getCellObject(x, y);

            // Document is being loaded and this cell already has a cached
            // answer from a previous evaluation: reuse it and skip the API.
            if (w.parent.building === true && cell.cache) {
                cell.v = cell.cache;
                return cell.v;
            }

            // Execute method
            try {
                // Only the server talks to the API. The browser falls through
                // and shows the placeholder until the server broadcast arrives.
                if (! isFront()) {
                    // Prompt
                    let prompt = Array.from(arguments);
                    // Query API
                    query.call({ w: w, x: x, y: y }, prompt, async function(result) {
                        // cell name
                        let cellName = w.helpers.getCellNameFromCoords(x, y);
                        // Save the answer in the local cache (also updates the
                        // cell value and persists with the document)
                        w.setCache(cellName, result);
                        // Same call, serialized for the peers. The client
                        // extension replays { method, worksheet, args } on
                        // every browser that has this document open.
                        const compressed = await Server.helpers.encodeZip({
                            method: 'setCache',
                            worksheet: w.options.worksheetId,
                            args: [cellName, result]
                        });
                        // Emit broadcast to the socket room of this document
                        Server.io.to(w.parent.guid).emit("JSS", compressed);
                    })
                }
            } catch (err) {
                console.log(err)
            }

            // Placeholder shown while the answer is pending
            return '...'
        }

        /**
         * =SUMMARY(...) is reserved and not implemented yet; it evaluates to
         * an empty result.
         */
        OPENAI.SUMMARY = function() {

        }

        return OPENAI;
    }

    /**
     * Configure the extension: create the OpenAI client and register the
     * optional hooks. Call it once on the server, with at least `apiKey`.
     *
     * The hooks are removed from `options` before the object is passed to the
     * OpenAI SDK constructor, every other key goes to the SDK untouched.
     *
     * @param {object} options - { apiKey, onbeforesend?, onsuccess?, ...OpenAI SDK options }
     */
    let Extension = function(options) {
        // Connect to the API
        if (options && options.apiKey) {
            // Register events
            if (typeof(options.onbeforesend) === 'function') {
                beforeSend = options.onbeforesend;
                delete options.onbeforesend;
            }
            if (typeof(options.onsuccess) === 'function') {
                success = options.onsuccess;
                delete options.onsuccess;
            }
            // Init openai
            instance = new OpenAI(options);
        }
    }

    /**
     * Extension hook called by jspreadsheet.setExtensions({ openai }, server)
     * with `this` = the jspreadsheet namespace.
     *
     * Registers the PROMPT formula on the formula engine. The formula
     * extension (formula-pro) must already be loaded, otherwise JSS.formula
     * is undefined and nothing is registered.
     *
     * @param {object} license - license information (unused here)
     * @param {object} server - server core from @jspreadsheet/server (undefined in the browser)
     */
    Extension.license = function(license, server) {
        if (JSS === null) {
            JSS = this;
        }

        // Server
        Server = server;

        if (JSS.formula) {
            let formulas = getFormulas();
            if (formulas) {
                JSS.formula.setFormula(formulas)
            }
        }
    }

    return Extension;

})));
