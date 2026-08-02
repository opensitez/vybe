// vybe-test: js/map_set_iterator_more_matrix/map_iterator_next_after_exhaustion_has_done_true
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

const it = new Map([["a", 1]]).keys();
it.next();
__check(__line(it.next().done), "true");
