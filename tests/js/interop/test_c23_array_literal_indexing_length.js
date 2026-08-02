// vybe-test: js/interop/test_c23_array_literal_indexing_length
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

let arr = [10, 20, 30, 40, 50];
        __check(__line(arr[0], arr[2], arr[4], arr.length), "10 30 50 5");
