// vybe-test: js/function_deep/arguments_length_reflects_call
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

function f(a, b, c) {
    return arguments.length;
}
__check(__line(f(1, 2)), "2");      // 2 args passed
__check(__line(f(1, 2, 3)), "3");   // 3 args passed
__check(__line(f(1, 2, 3, 4)), "4"); // 4 args passed
