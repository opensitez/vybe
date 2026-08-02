// vybe-test: js/promise_microtasks/catch_recovery_continues_chain
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

Promise.reject("bad")
    .catch(e => "recovered")
    .then(v => console.log(v));
