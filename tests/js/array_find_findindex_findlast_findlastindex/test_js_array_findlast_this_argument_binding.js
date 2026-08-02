// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_findlast_this_argument_binding
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

const ctx = { cap: 30 };
const nums = [10, 20, 40, 25];
__check(__line(nums.findLast(function(x) { return x < this.cap; }, ctx)), "25");
