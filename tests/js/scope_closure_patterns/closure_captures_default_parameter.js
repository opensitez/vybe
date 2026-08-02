// vybe-test: js/scope_closure_patterns/closure_captures_default_parameter
// origin: languages/js/tests/js/test_scope_closure_patterns.rs

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

function f(x = 10, g = () => x) {
    return g();
}
__check(__line(f()), "10");
__check(__line(f(20)), "20");
