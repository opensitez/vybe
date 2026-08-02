// vybe-test: js/bigint_bitwise_operations_masks/test_js_bigint_bitwise_and_or_xor_not
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

const a = 0b1100n, b = 0b1010n;
__check(__line(`${(a & b).toString()}:${(a | b).toString()}:${(a ^ b).toString()}:${(~a).toString()}`), "8:14:6:-13");
