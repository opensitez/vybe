// vybe-test: js/array_iteration_methods/reduce_empty_array_without_initial_throws_type_error
// origin: languages/js/tests/js/test_array_iteration_methods.rs

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

try {
    [].reduce((a, b) => a + b);
} catch (e) {
    __check(__line(e.name), "TypeError");
}
