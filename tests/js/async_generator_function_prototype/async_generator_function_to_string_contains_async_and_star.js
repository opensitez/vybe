// vybe-test: js/async_generator_function_prototype/async_generator_function_to_string_contains_async_and_star
// origin: languages/js/tests/js/test_async_generator_function_prototype.rs

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

async function* ping() { yield 1; } const text = Function.prototype.toString.call(ping); __check(__line(text.includes("async") && text.includes("*")), "true");
