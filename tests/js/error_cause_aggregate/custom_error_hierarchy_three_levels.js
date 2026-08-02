// vybe-test: js/error_cause_aggregate/custom_error_hierarchy_three_levels
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

class AppError extends Error { constructor(msg) { super(msg); this.name = "AppError"; } }
class NetworkError extends AppError { constructor(msg, code) { super(msg); this.name = "NetworkError"; this.code = code; } }
class TimeoutError extends NetworkError { constructor() { super("request timed out", 408); this.name = "TimeoutError"; } }
const e = new TimeoutError();
__check(__line(e instanceof TimeoutError), "true");
__check(__line(e instanceof NetworkError), "true");
__check(__line(e instanceof AppError), "true");
__check(__line(e instanceof Error), "true");
__check(__line(e.code), "408");
