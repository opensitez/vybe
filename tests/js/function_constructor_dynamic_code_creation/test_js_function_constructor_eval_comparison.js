// vybe-test: js/function_constructor_dynamic_code_creation/test_js_function_constructor_eval_comparison
// origin: languages/js/tests/js/test_js_function_constructor_dynamic_code_creation.rs

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

const fnEval = eval("(function(a) { return a * 3; })");
const fnConst = new Function("a", "return a * 3;");
__check(__line((fnEval(4) === fnConst(4)) + "|" + (fnEval.name !== fnConst.name)), "true|true");
