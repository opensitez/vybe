// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakmap_symbol_key_support_es2023
// origin: languages/js/tests/js/test_js_weakmap_weakset_object_key_lifecycle.rs

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

const wm = new WeakMap();
const sym = Symbol("weakKey");
wm.set(sym, "SymbolValue");
__check(__line(wm.get(sym) + "|" + wm.has(sym)), "SymbolValue|true");
