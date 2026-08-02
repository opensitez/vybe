// vybe-test: js/generator_async_delegate_errors/yield_star_on_array_iterates_elements
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

function* g(){yield*[10,20];} const a=[...g()]; __check(__line(a.join(",")), "10,20");
