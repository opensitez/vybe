// vybe-test: js/promise_advanced/async_await_rethrow_pattern
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

async function withRetry(fn) {
  try {
    return await fn();
  } catch (e) {
    return "fallback: " + e.message;
  }
}
async function main() {
  const r = await withRetry(() => { throw new Error("fail"); });
  console.log(r);
}
main();
