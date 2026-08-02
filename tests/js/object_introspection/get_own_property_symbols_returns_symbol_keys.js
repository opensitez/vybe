// vybe-test: js/object_introspection/get_own_property_symbols_returns_symbol_keys
// origin: languages/js/tests/js/test_object_introspection.rs

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

const sym = Symbol("tag");
const obj = { normal: 1, [sym]: "symbolValue" };
const syms = Object.getOwnPropertySymbols(obj);
__check(__line(syms.length), "1");
__check(__line(obj[syms[0]]), "symbolValue");
