// vybe-test: js/array_methods_new/tosorted_with_comparator
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

const arr = [10, 3, 7, 1];
const sorted = arr.toSorted((a, b) => b - a);
__check(__line(sorted.join(",")), "10,7,3,1");
