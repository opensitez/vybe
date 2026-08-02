// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_assignment_operators
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

let x = 0b1010;
x &= 0b1100;
__check(__line(x), "8");
x |= 0b0001;
__check(__line(x), "9");
x ^= 0b1001;
__check(__line(x), "0");
