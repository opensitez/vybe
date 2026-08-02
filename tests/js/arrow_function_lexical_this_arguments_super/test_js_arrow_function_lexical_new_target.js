// vybe-test: js/arrow_function_lexical_this_arguments_super/test_js_arrow_function_lexical_new_target
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

function Base() {
    const getNewTarget = () => new.target;
    return getNewTarget();
}
__check(__line(new Base() === Base), "true");
