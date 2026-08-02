// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_in_event_emitter_callback_simulation
// origin: languages/js/tests/js/test_js_async_arrow_functions_lexical_this.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

class Handler {
    constructor() { this.count = 0; }
    register(dispatcher) {
        dispatcher(async () => {
            this.count += 10;
        });
    }
}
let cb;
const h = new Handler();
h.register(fn => { cb = fn; });

cb().then(() => console.log(h.count));
