// vybe-test: js/map_set_edge_matrix/set_keys_matches_values_iterator
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

const s = new Set([1, 2]);
__check(__line(Array.from(s.keys()).join(",")), "1,2");
__check(__line(Array.from(s.values()).join(",")), "1,2");
