// vybe-test: js/function_arguments_hoisting/default_param_uses_previous_param
// origin: languages/js/tests/js/test_function_arguments_hoisting.rs

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

function greet(name, msg = "Hello " + name) {
    return msg;
}
__check(__line(greet("World")), "Hello World");
__check(__line(greet("World", "Hi World")), "Hi World");
