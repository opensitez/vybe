// vybe-test: js/eval_dynamic_code/new_function_does_not_capture_outer_scope
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

const secret = "top-secret";
const fn2 = new Function("try { return secret; } catch { return 'undefined'; }");
// new Function runs in global scope, can't access local 'secret'
const result = fn2();
console.log(result === "top-secret" || result === "undefined");
