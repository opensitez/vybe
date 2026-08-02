// vybe-test: js/function_invocation_matrix/bind_on_function_with_default_parameter_keeps_default_for_unbound_tail
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

function greet(greeting, name = "world") {
    return greeting + " " + name;
}
const hello = greet.bind(null, "hi");
__check(__line(hello()), "hi world");
