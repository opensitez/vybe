// vybe-test: js/exponentiation_operator_pow_precedence/test_js_exponentiation_null_and_undefined_operands
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

__check(__line((null ** 2) + "|" + (undefined ** 2) + "|" + (2 ** null)), "0|NaN|1");
