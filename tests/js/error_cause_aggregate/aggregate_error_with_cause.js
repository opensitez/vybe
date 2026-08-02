// vybe-test: js/error_cause_aggregate/aggregate_error_with_cause
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

const root = new Error("root");
const agg = new AggregateError([new Error("e1")], "agg", { cause: root });
__check(__line(agg.cause.message), "root");
