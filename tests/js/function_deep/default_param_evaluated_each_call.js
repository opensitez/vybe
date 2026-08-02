// vybe-test: js/function_deep/default_param_evaluated_each_call
// origin: languages/js/tests/js/test_function_deep.rs

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

let count = 0;
function f(x = ++count) { return x; }
__check(__line(f()), "1");    // 1
__check(__line(f()), "2");    // 2
__check(__line(f(99)), "99");  // 99 — no evaluation
__check(__line(count), "2");  // 2
