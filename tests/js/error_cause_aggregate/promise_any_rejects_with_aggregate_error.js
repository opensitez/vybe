// vybe-test: js/error_cause_aggregate/promise_any_rejects_with_aggregate_error
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

async function main() {
    try {
        await Promise.any([
            Promise.reject(new Error("a")),
            Promise.reject(new Error("b"))
        ]);
    } catch (e) {
        console.log(e instanceof AggregateError);
        console.log(e.errors.length);
        console.log(e.errors[0].message);
    }
}
main();
