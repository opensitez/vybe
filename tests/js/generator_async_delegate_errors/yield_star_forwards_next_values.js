// vybe-test: js/generator_async_delegate_errors/yield_star_forwards_next_values
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

function* inner(){yield 1;yield 2;} function* outer(){yield* inner();} const g=outer(); __check(__line(g.next().value), "1");__check(__line(g.next().value), "2");
