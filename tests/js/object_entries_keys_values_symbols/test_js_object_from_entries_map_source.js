// vybe-test: js/object_entries_keys_values_symbols/test_js_object_from_entries_map_source
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

const map = new Map([["key1", "val1"], ["key2", "val2"]]);
const obj = Object.fromEntries(map);
__check(__line(`${obj.key1}:${obj.key2}`), "val1:val2");
