// vybe-test: js/bigint_bitwise_operations_masks/test_js_bigint_bitwise_shift_assignment_operators
// origin: languages/js/tests/js/test_js_bigint_bitwise_operations_masks.rs

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

let x = 1n;
x <<= 4n;
__check(__line(x.toString()), "16");
x >>= 2n;
__check(__line(x.toString()), "4");
