// vybe-test: js/map_set_edge_matrix/map_string_and_number_keys_are_distinct
// origin: languages/js/tests/js/test_map_set_edge_matrix.rs

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

const m = new Map();
m.set("1", "string");
m.set(1, "number");
__check(__line(m.size), "2");
__check(__line(m.get("1")), "string");
__check(__line(m.get(1)), "number");
