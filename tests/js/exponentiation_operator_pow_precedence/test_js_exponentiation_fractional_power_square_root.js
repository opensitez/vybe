// vybe-test: js/exponentiation_operator_pow_precedence/test_js_exponentiation_fractional_power_square_root
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

__check(__line((16 ** 0.5) + "|" + (27 ** (1/3)).toFixed(1)), "4|3.0");
