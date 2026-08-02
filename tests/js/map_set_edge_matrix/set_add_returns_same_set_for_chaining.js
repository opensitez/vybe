// vybe-test: js/map_set_edge_matrix/set_add_returns_same_set_for_chaining
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

const s = new Set();
__check(__line(s.add(1) === s), "true");
