// vybe-test: js/error_types/rethrow_pattern
// origin: languages/js/tests/js/test_error_types.rs

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

function riskyOp() {
    throw new TypeError("wrong type");
}
try {
    try {
        riskyOp();
    } catch (e) {
        if (e instanceof TypeError) {
            throw e;
        }
        console.log("handled");
    }
} catch (e) {
    console.log("rethrown: " + e.message);
}
