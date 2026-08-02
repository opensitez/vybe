// vybe-test: js/ecma_functions/missing_arguments_produce_undefined
// origin: languages/js/tests/js/test_ecma_functions.rs

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

function show(a, b) {
    __check(__line(a), "1");
    __check(__line(b), "undefined");
}
show(1);
