// vybe-test: js/error_handling_advanced/custom_error_hierarchy
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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
class NetworkError extends AppError {
    constructor(message, statusCode) {
        super(message, "NETWORK");
        this.name = "NetworkError";
        this.statusCode = statusCode;
    }
}
const e = new NetworkError("Not Found", 404);
__check(__line(e instanceof Error), "true");
__check(__line(e instanceof AppError), "true");
__check(__line(e instanceof NetworkError), "true");
__check(__line(e.code), "NETWORK");
__check(__line(e.statusCode), "404");
__check(__line(e.message), "Not Found");
