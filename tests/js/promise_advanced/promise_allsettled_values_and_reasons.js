// vybe-test: js/promise_advanced/promise_allsettled_values_and_reasons
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

Promise.allSettled([
  Promise.resolve(42),
  Promise.reject("oops")
]).then(results => {
  console.log(results[0].value);
  console.log(results[1].reason);
});
