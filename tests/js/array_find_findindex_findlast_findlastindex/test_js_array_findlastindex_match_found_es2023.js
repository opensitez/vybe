// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_findlastindex_match_found_es2023
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

const numbers = [5, 12, 50, 130, 44];
const lastIdx = numbers.findLastIndex(element => element > 10);
__check(__line(lastIdx), "4");
