// vybe-test: js/function_deep/arguments_vs_rest_params
// origin: languages/js/tests/js/test_function_deep.rs

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

function withArgs() { return arguments[0]; }
const withRest = (...args) => args[0];
__check(__line(withArgs(42)), "42");
__check(__line(withRest(42)), "42");
