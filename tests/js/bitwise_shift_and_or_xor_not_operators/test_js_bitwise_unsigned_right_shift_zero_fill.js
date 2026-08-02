// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_unsigned_right_shift_zero_fill
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

__check(__line(`${-1 >>> 0}:${-5 >>> 2}`), "4294967295:1073741822");
