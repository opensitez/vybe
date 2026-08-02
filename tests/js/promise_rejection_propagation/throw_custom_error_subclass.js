// vybe-test: js/promise_rejection_propagation/throw_custom_error_subclass
// origin: languages/js/tests/js/test_promise_rejection_propagation.rs

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

class MyErr extends Error{}Promise.resolve(0).then(()=>{throw new MyErr("custom");}).catch(e=>console.log(e instanceof MyErr));
