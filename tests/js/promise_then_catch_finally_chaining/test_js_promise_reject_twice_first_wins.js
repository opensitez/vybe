// vybe-test: js/promise_then_catch_finally_chaining/test_js_promise_reject_twice_first_wins
// origin: languages/js/tests/js/test_js_promise_then_catch_finally_chaining.rs

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

new Promise((resolve, reject) => {
    reject("FirstErr");
    resolve("FirstVal");
})
.catch(err => console.log(err));
