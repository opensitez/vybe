// vybe-test: js/async_generator_function_prototype/async_generator_function_dynamic_constructor
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

const AsyncGenFn = Object.getPrototypeOf(async function*(){}).constructor; const f = new AsyncGenFn("a", "yield a * 2;"); __check(__line(f instanceof AsyncGenFn), "true");
