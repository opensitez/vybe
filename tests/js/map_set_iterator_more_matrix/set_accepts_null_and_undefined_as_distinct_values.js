// vybe-test: js/map_set_iterator_more_matrix/set_accepts_null_and_undefined_as_distinct_values
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

const s = new Set([null, undefined]);
__check(__line(s.size), "2");
__check(__line(s.has(null)), "true");
__check(__line(s.has(undefined)), "true");
