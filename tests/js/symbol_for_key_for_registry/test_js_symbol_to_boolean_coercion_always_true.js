// vybe-test: js/symbol_for_key_for_registry/test_js_symbol_to_boolean_coercion_always_true
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

const s1 = Symbol("");
const s2 = Symbol.for("test");
console.log(Boolean(s1) + "|" + Boolean(s2));
