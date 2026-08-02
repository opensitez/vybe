// vybe-test: js/ecma_operators/nullish_coalescing_preserves_falsey_non_nullish_values
// origin: languages/js/tests/js/test_ecma_operators.rs

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

__check(__line(false ?? true), "false");
__check(__line(0 ?? 10), "0");
__check(__line("" ?? "fallback"), "");
