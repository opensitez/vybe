// vybe-test: js/error_hierarchy/custom_error_catch_by_type
// origin: languages/js/tests/js/test_error_hierarchy.rs

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
    constructor(field, msg) {
        super(msg);
        this.name = "ValidationError";
        this.field = field;
    }
}
class AuthError extends Error {
    constructor(msg) { super(msg); this.name = "AuthError"; }
}

function handle(e) {
    if (e instanceof ValidationError) return "validation:" + e.field;
    if (e instanceof AuthError) return "auth:" + e.message;
    return "unknown";
}

__check(__line(handle(new ValidationError("email", "invalid"))), "validation:email");
__check(__line(handle(new AuthError("forbidden"))), "auth:forbidden");
__check(__line(handle(new Error("generic"))), "unknown");
