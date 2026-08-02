// vybe-test: js/bigint_bitwise_operations_masks/test_js_bigint_bitwise_unary_plus_throws_typeerror
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
    eval("+10n;");
} catch (e) {
    __check(__line("Unary Plus BigInt TypeError"), "Unary Plus BigInt TypeError");
}
