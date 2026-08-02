// vybe-test: js/generator_async_delegate_errors/generator_return_closes_iterator
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

function* g(){yield 1;yield 2;} const gen=g(); gen.next(); __check(__line(gen.return(9).value), "9");__check(__line(gen.next().done), "true");
