// vybe-test: js/error_handling_patterns/multiple_catch_types
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

class NetworkError extends Error { constructor(m) { super(m); this.name = "NetworkError"; } }
class TimeoutError extends NetworkError { constructor(m) { super(m); this.name = "TimeoutError"; } }
function handle(err) {
    if (err instanceof TimeoutError) return "timeout";
    if (err instanceof NetworkError) return "network";
    if (err instanceof Error) return "error";
    return "unknown";
}
__check(__line(handle(new TimeoutError("slow"))), "timeout");
__check(__line(handle(new NetworkError("down"))), "network");
__check(__line(handle(new Error("oops"))), "error");
