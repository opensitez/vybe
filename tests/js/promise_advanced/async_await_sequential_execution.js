// vybe-test: js/promise_advanced/async_await_sequential_execution
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

async function step(n) { return n * 2; }
async function main() {
  const a = await step(1);
  const b = await step(a);
  const c = await step(b);
  console.log(c);
}
main();
