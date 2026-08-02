// vybe-test: js/function_invocation_matrix/arguments_object_survives_method_call_in_body
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

function f(a) {
    __check(__line(arguments[0]), "ok");
    __check(__line("x".toUpperCase()), "X");
    __check(__line(arguments[0]), "ok");
}
f("ok");
