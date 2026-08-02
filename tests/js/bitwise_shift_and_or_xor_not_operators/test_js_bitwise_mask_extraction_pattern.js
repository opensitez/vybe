// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_mask_extraction_pattern
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

const flags = 0b1010;
const READ_FLAG = 0b0010;
const WRITE_FLAG = 0b0100;
__check(__line(`${(flags & READ_FLAG) !== 0}:${(flags & WRITE_FLAG) !== 0}`), "true:false");
