// vybe-test: js/map_set_edge_matrix/set_negative_zero_and_positive_zero_are_same_value
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
s.add(-0);
__check(__line(s.has(0)), "true");
__check(__line(s.size), "1");
