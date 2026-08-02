// vybe-test: js/misc_es_features/short_circuit_and_returns_falsy
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

__check(__line(0 && "never"), "0");
__check(__line(null && "never"), "null");
__check(__line(false && "never"), "false");
