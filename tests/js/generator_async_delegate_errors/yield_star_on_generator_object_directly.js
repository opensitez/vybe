// vybe-test: js/generator_async_delegate_errors/yield_star_on_generator_object_directly
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

function* inner(){yield "x";} function* outer(){yield* inner();} __check(__line(outer().next().value), "x");
