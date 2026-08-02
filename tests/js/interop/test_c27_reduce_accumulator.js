// vybe-test: js/interop/test_c27_reduce_accumulator
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

let result = [1, 2, 3, 4, 5].reduce((acc, x) => acc + x, 0);
        __check(__line(result), "15");
