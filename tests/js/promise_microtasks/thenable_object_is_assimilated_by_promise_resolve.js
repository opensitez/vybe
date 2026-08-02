// vybe-test: js/promise_microtasks/thenable_object_is_assimilated_by_promise_resolve
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

const thenable = {
    then(resolve) { resolve(99); }
};
Promise.resolve(thenable).then(v => console.log(v));
