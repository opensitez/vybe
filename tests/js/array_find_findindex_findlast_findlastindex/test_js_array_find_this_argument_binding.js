// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_find_this_argument_binding
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

const ctx = { threshold: 20 };
const nums = [10, 25, 30];
const res = nums.find(function(x) { return x > this.threshold; }, ctx);
__check(__line(res), "25");
