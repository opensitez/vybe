// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_and_or_xor_not_operators
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

const a = 0b1100, b = 0b1010;
__check(__line(`${a & b}:${a | b}:${a ^ b}:${(~a & 0xF)}`), "8:14:6:3");
