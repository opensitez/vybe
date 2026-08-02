// vybe-test: js/error_types_deep/aggregate_error_errors_array
// origin: languages/js/tests/js/test_error_types_deep.rs

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

const agg = new AggregateError([new Error("a"), new Error("b")], "multiple errors");
__check(__line(agg.message), "multiple errors");
__check(__line(agg.errors.length), "2");
__check(__line(agg.errors[0].message), "a");
