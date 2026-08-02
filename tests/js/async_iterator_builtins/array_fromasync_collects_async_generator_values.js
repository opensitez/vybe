// vybe-test: js/async_iterator_builtins/array_fromasync_collects_async_generator_values
// origin: languages/js/tests/js/test_async_iterator_builtins.rs

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

async function* numbers() {
  yield 1;
  yield 2;
  yield 3;
}
const result = await Array.fromAsync(numbers());
console.log(result.join(","));
