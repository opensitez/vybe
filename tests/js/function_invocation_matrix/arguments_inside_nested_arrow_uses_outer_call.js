// vybe-test: js/function_invocation_matrix/arguments_inside_nested_arrow_uses_outer_call
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

function outer() {
    const inner = () => arguments[1];
    __check(__line(inner()), "b");
}
outer("a", "b");
