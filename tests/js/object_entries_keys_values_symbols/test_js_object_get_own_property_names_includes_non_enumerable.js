// vybe-test: js/object_entries_keys_values_symbols/test_js_object_get_own_property_names_includes_non_enumerable
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

const obj = { a: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false });
__check(__line(Object.getOwnPropertyNames(obj).join(",")), "a,hidden");
