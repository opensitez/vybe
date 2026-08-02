// vybe-test: js/map_set_iterator_more_matrix/map_accepts_undefined_and_null_as_distinct_keys
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

const m = new Map([[undefined, 1], [null, 2]]);
__check(__line(m.size), "2");
__check(__line(m.get(undefined)), "1");
__check(__line(m.get(null)), "2");
