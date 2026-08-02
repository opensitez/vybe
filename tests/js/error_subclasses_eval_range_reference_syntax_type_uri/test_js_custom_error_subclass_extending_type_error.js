// vybe-test: js/error_subclasses_eval_range_reference_syntax_type_uri/test_js_custom_error_subclass_extending_type_error
// origin: languages/js/tests/js/test_js_error_subclasses_eval_range_reference_syntax_type_uri.rs

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

class ValidationError extends TypeError {
    constructor(field, message) {
        super(`${field}: ${message}`);
        this.name = "ValidationError";
        this.field = field;
    }
}
const err = new ValidationError("email", "Invalid format");
__check(__line(err.name + "|" + err.field + "|" + err.message + "|isTypeErr=" + (err instanceof TypeError)), "ValidationError|email|email: Invalid format|isTypeErr=true");
