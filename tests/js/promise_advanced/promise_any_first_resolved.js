// vybe-test: js/promise_advanced/promise_any_first_resolved
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

Promise.any([
  Promise.reject("err1"),
  Promise.resolve("ok"),
  Promise.resolve("second")
]).then(v => console.log(v));
