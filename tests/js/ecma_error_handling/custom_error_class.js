// vybe-test: js/ecma_error_handling/custom_error_class
// origin: languages/js/tests/js/test_ecma_error_handling.rs

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

class ValidationError extends Error {
    constructor(message, field) {
        super(message);
        this.field = field;
    }
}
try {
    throw new ValidationError("Required", "name");
} catch (e) {
    __check(__line(e.message), "Required");
    __check(__line(e.field), "name");
}
