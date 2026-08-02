// vybe-test: js/map_set_iterator_more_matrix/set_delete_then_add_reinserts_at_end
// origin: languages/js/tests/js/test_map_set_iterator_more_matrix.rs

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

const s = new Set([1, 2, 3]);
s.delete(2);
s.add(2);
__check(__line(Array.from(s).join(",")), "1,3,2");
