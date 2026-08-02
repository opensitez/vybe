// vybe-test: js/bigint_bitwise_operations_masks/test_js_bigint_unsigned_right_shift_prohibited
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

try {
    eval("1n >>> 2n;");
} catch (e) {
    __check(__line("BigInt Unsigned Right Shift TypeError"), "BigInt Unsigned Right Shift TypeError");
}
