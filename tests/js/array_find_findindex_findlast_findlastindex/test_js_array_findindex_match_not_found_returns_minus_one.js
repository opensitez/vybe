// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_findindex_match_not_found_returns_minus_one
// origin: languages/js/tests/js/test_js_array_find_findindex_findlast_findlastindex.rs

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

const numbers = [5, 2, 8];
const idx = numbers.findIndex(element => element > 10);
__check(__line(idx), "-1");
