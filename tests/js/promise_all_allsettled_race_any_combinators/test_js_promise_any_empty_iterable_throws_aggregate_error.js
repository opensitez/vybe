// vybe-test: js/promise_all_allsettled_race_any_combinators/test_js_promise_any_empty_iterable_throws_aggregate_error
// origin: languages/js/tests/js/test_js_promise_all_allsettled_race_any_combinators.rs

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

Promise.any([]).catch(err => {
    console.log((err instanceof AggregateError) + "|Count=" + err.errors.length);
});
