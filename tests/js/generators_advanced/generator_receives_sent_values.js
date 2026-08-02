// vybe-test: js/generators_advanced/generator_receives_sent_values
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

function* adder() {
  let sum = 0;
  while (true) {
    const n = yield sum;
    if (n === null) break;
    sum += n;
  }
  return sum;
}
const g = adder();
g.next();
g.next(5);
g.next(3);
const { value } = g.next(null);
console.log(value);
