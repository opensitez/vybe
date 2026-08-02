// vybe-test: js/error_types/custom_error_hierarchy
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

class HttpError extends Error {
    constructor(status, msg) {
        super(msg);
        this.name = "HttpError";
        this.status = status;
    }
}
class NotFoundError extends HttpError {
    constructor(resource) {
        super(404, resource + " not found");
        this.name = "NotFoundError";
    }
}
try {
    throw new NotFoundError("User");
} catch (e) {
    __check(__line(e.name), "NotFoundError");
    __check(__line(e.status), "404");
    __check(__line(e.message), "User not found");
    __check(__line(e instanceof NotFoundError), "true");
    __check(__line(e instanceof HttpError), "true");
    __check(__line(e instanceof Error), "true");
}
