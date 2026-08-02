// vybe-test: js/structured_clone_error_types_serialization/test_js_structured_clone_aggregate_error
// origin: languages/js/tests/js/test_js_structured_clone_error_types_serialization.rs

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

const err1 = new Error("Err1");
const err2 = new Error("Err2");
const agg = new AggregateError([err1, err2], "Bulk Failure");
const clone = structuredClone(agg);

__check(__line((clone instanceof AggregateError) + "|" + (clone.message === "Bulk Failure") + "|" + clone.errors.map(e => e.message).join(",")), "true|true|Err1,Err2");
