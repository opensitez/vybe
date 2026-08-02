// vybe-test: js/function_edge_cases/arrow_function_has_no_arguments_object
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

function outer() {
    const inner = () => arguments[0];
    return inner();
}
__check(__line(outer(42)), "42");
