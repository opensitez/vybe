// vybe-test: js/generator_async_delegate_errors/generator_yield_in_try_finally_preserves_flow
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

function* g(){try{yield 1;}finally{yield 2;}} __check(__line([...g()].join(",")), "1,2");
