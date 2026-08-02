// vybe-test: js/map_set_iteration_entries_keys_values/test_js_map_for_of_iteration_yields_entry_tuples
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

const map = new Map([["k1", 10], ["k2", 20]]);
const res = [];
for (const [k, v] of map) {
    res.push(`${k}:${v}`);
}
console.log(res.join(","));
