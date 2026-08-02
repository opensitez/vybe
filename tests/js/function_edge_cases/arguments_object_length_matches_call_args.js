// vybe-test: js/function_edge_cases/arguments_object_length_matches_call_args
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

function f() { return arguments.length; }
__check(__line(f(1, 2, 3)), "3");
__check(__line(f()), "0");
