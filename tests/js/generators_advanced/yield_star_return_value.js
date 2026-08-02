// vybe-test: js/generators_advanced/yield_star_return_value
// origin: languages/js/tests/js/test_generators_advanced.rs

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

function* inner() { yield 1; return "done"; }
function* outer() {
  const result = yield* inner();
  yield result;
}
__check(__line([...outer()].join(",")), "1,done");
