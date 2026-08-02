// vybe-test: js/error_types_deep/eval_error_exists
// origin: languages/js/tests/js/test_error_types_deep.rs

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

const e = new EvalError("test");
__check(__line(e instanceof EvalError), "true");
__check(__line(e instanceof Error), "true");
__check(__line(e.name), "EvalError");
