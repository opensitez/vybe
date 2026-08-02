// vybe-test: js/function_invocation_matrix/arguments_is_empty_when_no_values_passed
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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

function f() {
    __check(__line(arguments.length), "0");
    __check(__line(arguments[0]), "undefined");
}
f();
