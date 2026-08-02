// vybe-test: js/closure_scope_deep_patterns/closure_captures_outer_arguments
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

function outer() {
    return () => arguments[0] + arguments[1];
}
const f = outer(10, 20);
__check(__line(f()), "30");
