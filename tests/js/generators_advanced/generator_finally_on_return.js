// vybe-test: js/generators_advanced/generator_finally_on_return
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

const steps = [];
function* gen() {
  try {
    yield 1;
    yield 2;
  } finally {
    steps.push("finally");
  }
}
const g = gen();
g.next();
g.return("end");
__check(__line(steps.join(",")), "finally");
