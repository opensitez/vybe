// vybe-test: js/structured_clone_error_types_serialization/test_js_structured_clone_eval_error_subclass
// origin: languages/js/tests/js/test_js_structured_clone_error_types_serialization.rs

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

const err = new EvalError("Eval error");
const clone = structuredClone(err);
__check(__line((clone instanceof EvalError) + "|" + clone.message), "true|Eval error");
