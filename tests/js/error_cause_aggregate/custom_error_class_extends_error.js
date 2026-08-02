// vybe-test: js/error_cause_aggregate/custom_error_class_extends_error
// origin: languages/js/tests/js/test_error_cause_aggregate.rs

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
        this.name = "ValidationError";
        this.field = field;
    }
}
const e = new ValidationError("required", "email");
__check(__line(e instanceof ValidationError), "true");
__check(__line(e instanceof Error), "true");
__check(__line(e.name), "ValidationError");
__check(__line(e.field), "email");
__check(__line(e.message), "required");
