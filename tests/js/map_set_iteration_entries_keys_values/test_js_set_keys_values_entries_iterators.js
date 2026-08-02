// vybe-test: js/map_set_iteration_entries_keys_values/test_js_set_keys_values_entries_iterators
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

const set = new Set(["x", "y"]);
__check(__line([...set.keys()].join(",")), "x,y");
__check(__line([...set.values()].join(",")), "x,y");
__check(__line([...set.entries()].map(e => e.join("=")).join(",")), "x=x,y=y");
