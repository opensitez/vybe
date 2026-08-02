// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_recursion_with_outer_environment
// origin: languages/js/tests/js/test_js_scope_closures_lexical_environment_capture.rs

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

function makeFactorialMemo() {
    const memo = {};
    function fact(n) {
        if (n <= 1) return 1;
        if (memo[n]) return memo[n];
        return (memo[n] = n * fact(n - 1));
    }
    return fact;
}
const fact = makeFactorialMemo();
__check(__line(fact(5) + "|" + fact(5)), "120|120");
