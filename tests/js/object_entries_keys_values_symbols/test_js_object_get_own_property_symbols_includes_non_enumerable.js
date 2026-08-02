// vybe-test: js/object_entries_keys_values_symbols/test_js_object_get_own_property_symbols_includes_non_enumerable
// origin: languages/js/tests/js/test_js_object_entries_keys_values_symbols.rs

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

const s = Symbol("hiddenSym");
const obj = {};
Object.defineProperty(obj, s, { value: "secret", enumerable: false });
const symbols = Object.getOwnPropertySymbols(obj);
__check(__line(symbols.length + "|" + obj[symbols[0]]), "1|secret");
