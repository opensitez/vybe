// vybe-test: js/eval_dynamic_code/eval_does_not_leak_let_to_outer
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

eval("let innerLet = 5;");
let threw = false;
try { innerLet; } catch { threw = true; }
__check(__line(threw), "true");
