// vybe-test: js/error_aggregate_error_cause_property/test_js_aggregate_error_errors_array_is_copied
// origin: languages/js/tests/js/test_js_error_aggregate_error_cause_property.rs

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

const input = [1, 2];
const agg = new AggregateError(input, "Msg");
input.push(3); // Modifying input array after construction
__check(__line(agg.errors.length), "2"); // agg.errors is frozen / copied snapshot
