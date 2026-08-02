// vybe-test: js/map_set_iterator_more_matrix/set_default_iterator_matches_values
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

const a = Array.from(new Set([1, 2]).values()).join(",");
const b = Array.from(new Set([1, 2])[Symbol.iterator]()).join(",");
__check(__line(a === b), "true");
