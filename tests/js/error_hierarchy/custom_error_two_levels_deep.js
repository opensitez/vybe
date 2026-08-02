// vybe-test: js/error_hierarchy/custom_error_two_levels_deep
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

class BaseError extends Error {
    constructor(msg) { super(msg); this.name = "BaseError"; }
}
class NetworkError extends BaseError {
    constructor(msg, status) {
        super(msg);
        this.name = "NetworkError";
        this.status = status;
    }
}
const e = new NetworkError("timeout", 503);
__check(__line(e instanceof NetworkError), "true");
__check(__line(e instanceof BaseError), "true");
__check(__line(e instanceof Error), "true");
__check(__line(e.status), "503");
