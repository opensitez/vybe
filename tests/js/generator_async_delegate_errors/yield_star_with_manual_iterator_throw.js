// vybe-test: js/generator_async_delegate_errors/yield_star_with_manual_iterator_throw
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

const it={*[Symbol.iterator](){yield 1;throw new Error("it");}}; function* g(){try{yield* it;}catch(e){yield "caught";}} __check(__line([...g()].join(",")), "1,caught");
