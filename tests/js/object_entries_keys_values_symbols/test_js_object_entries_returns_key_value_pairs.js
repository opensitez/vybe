// vybe-test: js/object_entries_keys_values_symbols/test_js_object_entries_returns_key_value_pairs
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

const obj = { x: 1, y: 2 };
const pairs = Object.entries(obj);
__check(__line(pairs.map(p => p.join("=")).join("|")), "x=1|y=2");
