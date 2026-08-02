// vybe-test: js/promise_microtasks/resolve_with_promise_adopts_state
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

const inner = new Promise(resolve => setTimeout(() => resolve("inner"), 0));
Promise.resolve(inner).then(v => console.log(v));
