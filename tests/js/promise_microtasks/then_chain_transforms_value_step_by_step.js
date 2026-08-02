// vybe-test: js/promise_microtasks/then_chain_transforms_value_step_by_step
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

Promise.resolve(5)
    .then(n => n * 2)
    .then(n => n + 3)
    .then(n => n.toString())
    .then(s => console.log(s));
