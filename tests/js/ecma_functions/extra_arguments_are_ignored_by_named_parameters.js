// vybe-test: js/ecma_functions/extra_arguments_are_ignored_by_named_parameters
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

function add(a, b) {
    __check(__line(a + b), "5");
}
add(2, 3, 4, 5);
