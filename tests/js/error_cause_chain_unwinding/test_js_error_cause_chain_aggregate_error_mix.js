// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_chain_aggregate_error_mix
// origin: languages/js/tests/js/test_js_error_cause_chain_unwinding.rs

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

const e1 = new Error("SubTaskA Failed");
const e2 = new Error("SubTaskB Failed");
const agg = new AggregateError([e1, e2], "Batch Failed");
const root = new Error("Job Failed", { cause: agg });

__check(__line(root.cause.errors.map(e => e.message).join(",")), "SubTaskA Failed,SubTaskB Failed");
