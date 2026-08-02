// vybe-test: js/error_types/custom_error_class
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

class AppError extends Error {
    constructor(message, code) {
        super(message);
        this.name = "AppError";
        this.code = code;
    }
}
try {
    throw new AppError("not found", 404);
} catch (e) {
    __check(__line(e.name), "AppError");
    __check(__line(e.message), "not found");
    __check(__line(e.code), "404");
    __check(__line(e instanceof AppError), "true");
    __check(__line(e instanceof Error), "true");
}
