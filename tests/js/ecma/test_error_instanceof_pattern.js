// vybe-test: js/ecma/test_error_instanceof_pattern
// origin: languages/js/tests/js/js_ecma_test.rs

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

function validate(x) {
            if (x < 0) throw new RangeError("must be positive");
            return x;
        }
        try {
            validate(-1);
        } catch (e) {
            __check(__line(e.name + ": " + e.message), "RangeError: must be positive");
        }
