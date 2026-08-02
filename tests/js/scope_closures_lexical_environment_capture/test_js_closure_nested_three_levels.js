// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_nested_three_levels
// origin: languages/js/tests/js/test_js_scope_closures_lexical_environment_capture.rs

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

function level1(a) {
    return function level2(b) {
        return function level3(c) {
            return a + b + c;
        };
    };
}
__check(__line(level1(10)(20)(30)), "60");
