// vybe-test: js/map_set_edge_matrix/map_delete_present_returns_true
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

const m = new Map([["x", 1]]);
__check(__line(m.delete("x")), "true");
__check(__line(m.has("x")), "false");
