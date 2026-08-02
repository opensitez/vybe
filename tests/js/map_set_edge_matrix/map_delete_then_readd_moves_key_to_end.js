// vybe-test: js/map_set_edge_matrix/map_delete_then_readd_moves_key_to_end
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

const m = new Map([["a", 1], ["b", 2]]);
m.delete("a");
m.set("a", 3);
__check(__line(Array.from(m.keys()).join(",")), "b,a");
