// vybe-test: js/ecma_arrays/array_at
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

const arr = [10, 20, 30, 40, 50];
__check(__line(arr.at(0)), "10");
__check(__line(arr.at(-1)), "50");
__check(__line(arr.at(-2)), "40");
