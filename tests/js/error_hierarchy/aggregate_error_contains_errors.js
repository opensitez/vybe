// vybe-test: js/error_hierarchy/aggregate_error_contains_errors
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

const errors = [new Error("e1"), new TypeError("e2")];
const agg = new AggregateError(errors, "multiple failures");
__check(__line(agg instanceof AggregateError), "true");
__check(__line(agg instanceof Error), "true");
__check(__line(agg.message), "multiple failures");
__check(__line(agg.errors.length), "2");
__check(__line(agg.errors[0].message), "e1");
