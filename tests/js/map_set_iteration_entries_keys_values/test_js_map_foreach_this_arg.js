// vybe-test: js/map_set_iteration_entries_keys_values/test_js_map_foreach_this_arg
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

const context = { prefix: "Item" };
const map = new Map([["1", "A"]]);
map.forEach(function(val, key) {
    console.log(`${this.prefix}:${key}->${val}`);
}, context);
