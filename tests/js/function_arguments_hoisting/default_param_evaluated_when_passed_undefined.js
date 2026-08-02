// vybe-test: js/function_arguments_hoisting/default_param_evaluated_when_passed_undefined
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

function f(x = "default") { return x; }
__check(__line(f(undefined)), "default");
__check(__line(f(null)), "null");
