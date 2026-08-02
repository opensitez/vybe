// vybe-test: js/async_function_prototype/async_function_bind_partial_preserves_async_prototype
// origin: languages/js/tests/js/test_async_function_prototype.rs

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

async function add(a, b) { return a + b; } const plusOne = add.bind(null, 1); __check(__line(Object.getPrototypeOf(plusOne) === AsyncFunction.prototype), "true");
