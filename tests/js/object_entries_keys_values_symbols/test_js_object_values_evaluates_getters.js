// vybe-test: js/object_entries_keys_values_symbols/test_js_object_values_evaluates_getters
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

const obj = {
    get a() { return "GetterA"; }
};
__check(__line(Object.values(obj).join(",")), "GetterA");
