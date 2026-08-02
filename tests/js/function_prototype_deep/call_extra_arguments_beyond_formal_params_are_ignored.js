// vybe-test: js/function_prototype_deep/call_extra_arguments_beyond_formal_params_are_ignored
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

function takeOne(a) { return a; } __check(__line(takeOne.call(null, 1, 2, 3)), "1");
