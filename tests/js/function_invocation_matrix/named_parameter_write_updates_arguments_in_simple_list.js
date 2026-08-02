// vybe-test: js/function_invocation_matrix/named_parameter_write_updates_arguments_in_simple_list
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
    a = 8;
    __check(__line(arguments[0]), "8");
}
f(1);
