// vybe-test: js/object_entries_keys_values_symbols/test_js_object_get_own_property_symbols_reflects_symbols_only
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

const s1 = Symbol("s1");
const s2 = Symbol("s2");
const obj = { stringKey: 1, [s1]: "val1", [s2]: "val2" };
const symbols = Object.getOwnPropertySymbols(obj);
__check(__line(symbols.length + "|" + (symbols[0] === s1) + "|" + (symbols[1] === s2)), "2|true|true");
