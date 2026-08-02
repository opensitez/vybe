// vybe-test: js/promise_advanced/promise_finally_runs_on_resolve
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

const steps = [];
Promise.resolve("ok")
  .finally(() => steps.push("finally"))
  .then(v => { steps.push(v); console.log(steps.join(",")); });
