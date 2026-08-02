// vybe-test: js/misc_es_features/short_circuit_or_returns_first_truthy
// origin: languages/js/tests/js/test_misc_es_features.rs

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

__check(__line(0 || "fallback"), "fallback");
__check(__line("" || 42), "42");
__check(__line("first" || "second"), "first");
