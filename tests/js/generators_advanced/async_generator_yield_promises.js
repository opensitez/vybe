// vybe-test: js/generators_advanced/async_generator_yield_promises
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

async function* gen() {
  yield await Promise.resolve(1);
  yield await Promise.resolve(2);
}
async function run() {
  const result = [];
  for await (const v of gen()) result.push(v);
  console.log(result.join(","));
}
run();
