// vybe-test: js/function_edge_cases/default_param_can_call_function
// origin: languages/js/tests/js/test_function_edge_cases.rs

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

function getDefault() { return 42; }
function f(x = getDefault()) { return x; }
__check(__line(f()), "42");
__check(__line(f(0)), "0");
