// vybe-test: js/generator_async_delegate_errors/yield_star_throw_from_nested_delegate
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

function* a(){throw "a";} function* b(){yield* a();} function* c(){try{yield* b();}catch(e){yield e;}} __check(__line([...c()][0]), "a");
