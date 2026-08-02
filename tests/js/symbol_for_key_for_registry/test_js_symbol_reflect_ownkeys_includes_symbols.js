// vybe-test: js/symbol_for_key_for_registry/test_js_symbol_reflect_ownkeys_includes_symbols
// origin: languages/js/tests/js/test_js_symbol_for_key_for_registry.rs

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

const s = Symbol("s");
const obj = { a: 1, [s]: 2 };
const keys = Reflect.ownKeys(obj);
__check(__line(keys.length + "|" + (keys[1] === s)), "2|true");
