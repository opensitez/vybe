// vybe-test: js/promise_advanced/promise_chaining_async_values
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

async function main() {
  const result = await Promise.resolve(10)
    .then(async x => x + await Promise.resolve(5))
    .then(x => x * 2);
  console.log(result);
}
main();
