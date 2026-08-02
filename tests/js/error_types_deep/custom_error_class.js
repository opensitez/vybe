// vybe-test: js/error_types_deep/custom_error_class
// origin: languages/js/tests/js/test_error_types_deep.rs

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

class AppError extends Error {
    constructor(msg, code) {
        super(msg);
        this.name = "AppError";
        this.code = code;
    }
}
const e = new AppError("failed", 404);
__check(__line(e instanceof AppError), "true");
__check(__line(e instanceof Error), "true");
__check(__line(e.name), "AppError");
__check(__line(e.code), "404");
__check(__line(e.message), "failed");
