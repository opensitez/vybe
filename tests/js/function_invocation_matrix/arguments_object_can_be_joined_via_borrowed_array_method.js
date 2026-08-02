// vybe-test: js/function_invocation_matrix/arguments_object_can_be_joined_via_borrowed_array_method
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

function f() {
    __check(__line(Array.prototype.join.call(arguments, "-")), "a-b-c");
}
f("a", "b", "c");
