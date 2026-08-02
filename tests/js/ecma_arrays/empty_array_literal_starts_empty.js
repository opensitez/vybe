// vybe-test: js/ecma_arrays/empty_array_literal_starts_empty
// origin: languages/js/tests/js/test_ecma_arrays.rs

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

const arr = [];
__check(__line(arr.length), "0");
arr.push(1);
arr.push(2);
__check(__line(arr.join(",")), "1,2");
