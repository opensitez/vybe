// vybe-test: js/object_entries_keys_values_symbols/test_js_object_keys_integer_indices_sorted_first
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

const obj = { 10: "ten", 2: "two", "b": "B", 1: "one", "a": "A" };
__check(__line(Object.keys(obj).join(",")), "1,2,10,b,a");
