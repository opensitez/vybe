// vybe-test: js/type_coercion_deep/tostring_bigint_and_number_from_explicit_calls
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

__check(__line(String(10n)), "10");
__check(__line(Number(10n)), "10");
__check(__line(Number(1n + 2n)), "3");
