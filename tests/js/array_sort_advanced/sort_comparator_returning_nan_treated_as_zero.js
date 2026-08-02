// vybe-test: js/array_sort_advanced/sort_comparator_returning_nan_treated_as_zero
// origin: languages/js/tests/js/test_array_sort_advanced.rs

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

const arr = [3, 1, 2];
arr.sort(() => NaN);
__check(__line(arr.join(",")), "3,1,2");
