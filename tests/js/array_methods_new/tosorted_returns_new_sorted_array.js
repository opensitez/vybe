// vybe-test: js/array_methods_new/tosorted_returns_new_sorted_array
// origin: languages/js/tests/js/test_array_methods_new.rs

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

const arr = [3, 1, 4, 1, 5, 9];
const sorted = arr.toSorted();
__check(__line(sorted.join(",")), "1,1,3,4,5,9");
__check(__line(arr[0]), "3"); // original unchanged
