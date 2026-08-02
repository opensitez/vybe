// vybe-test: js/map_set_edge_matrix/set_delete_then_readd_moves_value_to_end
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

const s = new Set(["a", "b"]);
s.delete("a");
s.add("a");
__check(__line(Array.from(s).join(",")), "b,a");
