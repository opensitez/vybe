// vybe-test: js/generators_advanced/generator_throw_caught_inside
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
  try {
    yield 1;
    yield 2;
  } catch (e) {
    yield "caught: " + e;
  }
}
const g = gen();
__check(__line(g.next().value), "1");
__check(__line(g.throw("oops").value), "caught: oops");
__check(__line(g.next().done), "true");
