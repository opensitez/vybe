// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_shift_assignment_operators
// origin: languages/js/tests/js/test_js_bitwise_shift_and_or_xor_not_operators.rs

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

let a = 1;
a <<= 3;
__check(__line(a), "8");
a >>= 1;
__check(__line(a), "4");
a >>>= 1;
__check(__line(a), "2");
