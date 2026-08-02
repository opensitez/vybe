// vybe-test: js/string_array_advanced/array_sort_numbers_correct
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

let arr = [10, 1, 21, 2];
arr.sort((a, b) => a - b);
__check(__line(arr.join(",")), "1,2,10,21");
