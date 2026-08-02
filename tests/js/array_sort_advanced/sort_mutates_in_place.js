// vybe-test: js/array_sort_advanced/sort_mutates_in_place
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
const ret = arr.sort();
__check(__line(ret === arr), "true"); // same array
__check(__line(arr.join(",")), "1,2,3");
