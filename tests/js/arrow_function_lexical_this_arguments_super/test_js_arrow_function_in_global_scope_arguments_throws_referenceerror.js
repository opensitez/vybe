// vybe-test: js/arrow_function_lexical_this_arguments_super/test_js_arrow_function_in_global_scope_arguments_throws_referenceerror
// origin: languages/js/tests/js/test_js_arrow_function_lexical_this_arguments_super.rs

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

const arrow = () => typeof arguments;
__check(__line(arrow()), "undefined");
