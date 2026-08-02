// vybe-test: js/promise_advanced/promise_then_returns_new_promise
// origin: languages/js/tests/js/test_promise_advanced.rs

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

const p1 = Promise.resolve("a");
const p2 = p1.then(v => v + "b");
p2.then(v => console.log(v));
