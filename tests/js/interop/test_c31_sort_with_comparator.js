// vybe-test: js/interop/test_c31_sort_with_comparator
// origin: languages/js/tests/js/js_interop_test.rs

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

let arr = [5, 3, 8, 1, 9, 2];
        arr.sort((a, b) => a - b);
        __check(__line(arr), "1,2,3,5,8,9");
