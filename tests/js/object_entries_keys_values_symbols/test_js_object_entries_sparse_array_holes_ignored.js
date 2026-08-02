// vybe-test: js/object_entries_keys_values_symbols/test_js_object_entries_sparse_array_holes_ignored
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

const sparse = [10, , 30];
const entries = Object.entries(sparse);
__check(__line(entries.map(e => e.join("=")).join("|")), "0=10|2=30");
