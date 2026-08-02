// vybe-test: js/promise_advanced/promise_race_already_resolved_wins
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

Promise.race([
  Promise.resolve("first"),
  Promise.resolve("second"),
  Promise.resolve("third")
]).then(v => console.log(v));
