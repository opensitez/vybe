// vybe-test: js/generator_async_delegate_errors/generator_next_after_throw_still_closed
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

function* g(){throw new Error("e");} const gen=g(); try{gen.next();}catch{} __check(__line(gen.next().done), "true");
