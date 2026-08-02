// vybe-test: js/error_cause_aggregate/custom_error_with_cause_in_hierarchy
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

class ServiceError extends Error {
    constructor(message, options) {
        super(message, options);
        this.name = "ServiceError";
    }
}
const root = new Error("DB connection failed");
const svc = new ServiceError("Cannot fetch user", { cause: root });
__check(__line(svc.name), "ServiceError");
__check(__line(svc.cause.message), "DB connection failed");
