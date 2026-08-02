// vybe-test: js/map_set_iteration_entries_keys_values/test_js_set_foreach_callback_arguments
// origin: languages/js/tests/js/test_js_map_set_iteration_entries_keys_values.rs

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

const set = new Set(["elem"]);
set.forEach((val1, val2, s) => {
    console.log(`${val1}:${val2}|isSet=${s === set}`); // In Set forEach, first and second args are identical element!
});
