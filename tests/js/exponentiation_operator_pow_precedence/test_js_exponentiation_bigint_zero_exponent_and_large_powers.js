// vybe-test: js/exponentiation_operator_pow_precedence/test_js_exponentiation_bigint_zero_exponent_and_large_powers
// origin: languages/js/tests/js/test_js_exponentiation_operator_pow_precedence.rs

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

__check(__line(`${(0n ** 0n).toString()}:${(5n ** 0n).toString()}:${(2n ** 64n).toString()}`), "1:1:18446744073709551616");
