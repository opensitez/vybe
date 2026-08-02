// vybe-test: js/error_handling_patterns/validation_error_with_field
// origin: languages/js/tests/js/test_error_handling_patterns.rs

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
    constructor(field, message) {
        super(message);
        this.name = "ValidationError";
        this.field = field;
    }
}
function validateAge(age) {
    if (typeof age !== "number") throw new ValidationError("age", "must be a number");
    if (age < 0) throw new ValidationError("age", "must be non-negative");
    return age;
}
try {
    validateAge(-1);
} catch (e) {
    __check(__line(e instanceof ValidationError), "true");
    __check(__line(e.field), "age");
    __check(__line(e.message), "must be non-negative");
}
