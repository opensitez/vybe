// vybe-test: js/promise_advanced/async_await_promise_all_parallel
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

async function double(x) { return x * 2; }
async function main() {
  const [a, b, c] = await Promise.all([double(1), double(2), double(3)]);
  console.log(a, b, c);
}
main();
