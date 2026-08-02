// vybe-test: js/generator_async_delegate_errors/yield_star_throw_propagates_to_outer_iterator
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

function* inner(){yield 1;throw new Error("inner");} function* outer(){try{yield* inner();}catch(e){yield e.message;}} const g=outer(); g.next(); __check(__line(g.next().value), "inner");
