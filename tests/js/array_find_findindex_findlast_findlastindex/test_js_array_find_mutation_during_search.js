// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_find_mutation_during_search
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

const nums = [0, 1, 2];
const res = nums.find((x, idx, a) => {
    if (idx === 0) a[1] = 99; // Mutates index 1 before visited
    return x === 99;
});
__check(__line(res), "99");
