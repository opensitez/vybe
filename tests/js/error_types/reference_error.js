// vybe-test: js/error_types/reference_error
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

try {
    throw new ReferenceError("x is not defined");
} catch (e) {
    __check(__line(e instanceof ReferenceError), "true");
    __check(__line(e.message), "x is not defined");
}
