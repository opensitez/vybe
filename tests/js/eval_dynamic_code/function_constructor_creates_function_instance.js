// vybe-test: js/eval_dynamic_code/function_constructor_creates_function_instance
// origin: languages/js/tests/js/test_eval_dynamic_code.rs

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

const f = new Function("return 42");
__check(__line(f instanceof Function), "true");
__check(__line(typeof f), "function");
