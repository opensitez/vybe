// vybe-test: js/error_cause_aggregate/wrapping_unknown_errors_in_known_type
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

function safeDivide(a, b) {
    if (b === 0) throw new RangeError("division by zero");
    return a / b;
}
function compute(a, b) {
    try { return safeDivide(a, b); }
    catch (e) { throw new Error("compute failed", { cause: e }); }
}
let msg = "";
try { compute(10, 0); }
catch (e) { msg = e.message + ":" + e.cause.constructor.name; }
__check(__line(msg), "compute failed:RangeError");
