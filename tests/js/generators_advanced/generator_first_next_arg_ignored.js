// vybe-test: js/generators_advanced/generator_first_next_arg_ignored
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

function* gen() {
  const x = yield "first";
  yield x * 2;
}
const g = gen();
g.next("ignored");
const { value } = g.next(10);
__check(__line(value), "20");
