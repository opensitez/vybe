// vybe-test: js/generators_advanced/async_generator_basic
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

async function* asyncRange(start, end) {
  for (let i = start; i <= end; i++) yield i;
}
async function collect() {
  const result = [];
  for await (const v of asyncRange(1, 4)) result.push(v);
  console.log(result.join(","));
}
collect();
