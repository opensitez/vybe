// vybe-test: js/object_methods_deep/get_own_property_symbols_finds_symbol_keys
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const sym = Symbol("test");
const obj = { [sym]: "value", normal: 1 };
const symbols = Object.getOwnPropertySymbols(obj);
__check(__line(symbols.length), "1");
__check(__line(obj[symbols[0]]), "value");
