// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_findlast_stops_at_first_match_from_end
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

let calls = 0;
const nums = [1, 2, 3, 4, 5];
nums.findLast(x => {
    calls++;
    return x === 4;
});
__check(__line(calls), "2");
