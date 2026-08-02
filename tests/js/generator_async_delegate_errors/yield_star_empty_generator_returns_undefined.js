// vybe-test: js/generator_async_delegate_errors/yield_star_empty_generator_returns_undefined
// origin: languages/js/tests/js/test_generator_async_delegate_errors.rs

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

function* empty(){} function* g(){const r=yield* empty(); yield r===undefined;} __check(__line([...g()][0]), "true");
