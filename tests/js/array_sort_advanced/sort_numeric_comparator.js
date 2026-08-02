// vybe-test: js/array_sort_advanced/sort_numeric_comparator
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

const arr = [10, 9, 2, 21, 3];
arr.sort((a, b) => a - b);
__check(__line(arr.join(",")), "2,3,9,10,21");
