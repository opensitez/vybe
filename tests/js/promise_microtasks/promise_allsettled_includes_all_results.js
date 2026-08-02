// vybe-test: js/promise_microtasks/promise_allsettled_includes_all_results
// origin: languages/js/tests/js/test_promise_microtasks.rs

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

Promise.allSettled([
    Promise.resolve("ok"),
    Promise.reject("err"),
    Promise.resolve("ok2"),
]).then(results => {
    console.log(results[0].status + ":" + results[0].value);
    console.log(results[1].status + ":" + results[1].reason);
    console.log(results[2].status + ":" + results[2].value);
});
