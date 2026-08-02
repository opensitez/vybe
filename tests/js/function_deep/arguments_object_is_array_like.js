// vybe-test: js/function_deep/arguments_object_is_array_like
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

function f() {
    return Array.from(arguments).join(",");
}
__check(__line(f(1, 2, 3)), "1,2,3");
