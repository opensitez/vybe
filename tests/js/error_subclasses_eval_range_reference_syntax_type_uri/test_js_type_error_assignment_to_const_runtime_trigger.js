// vybe-test: js/error_subclasses_eval_range_reference_syntax_type_uri/test_js_type_error_assignment_to_const_runtime_trigger
// origin: languages/js/tests/js/test_js_error_subclasses_eval_range_reference_syntax_type_uri.rs

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

const c = 10;
try {
    eval("c = 20;");
} catch (e) {
    __check(__line(e.name + "|isType=" + (e instanceof TypeError)), "TypeError|isType=true");
}
