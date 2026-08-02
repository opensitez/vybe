// vybe-test: js/closure_scope_deep/closure_over_async_state
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

function createAsyncCounter() {
    let count = 0;
    return async function() {
        await Promise.resolve(); // yield control
        return ++count;
    };
}
const next = createAsyncCounter();
Promise.all([next(), next(), next()]).then(results => {
    console.log(results.join(","));
});
