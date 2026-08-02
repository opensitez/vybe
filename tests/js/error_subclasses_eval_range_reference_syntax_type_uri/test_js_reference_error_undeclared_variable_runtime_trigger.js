// vybe-test: js/error_subclasses_eval_range_reference_syntax_type_uri/test_js_reference_error_undeclared_variable_runtime_trigger
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

try {
    nonExistentVariableReference;
} catch (e) {
    __check(__line(e.name + "|" + (e instanceof ReferenceError)), "ReferenceError|true");
}
