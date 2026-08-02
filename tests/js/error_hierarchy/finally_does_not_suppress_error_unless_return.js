// vybe-test: js/error_hierarchy/finally_does_not_suppress_error_unless_return
// origin: languages/js/tests/js/test_error_hierarchy.rs

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

function f() {
    try { throw new Error("original"); }
    finally { /* no return, error propagates */ }
}
try { f(); } catch (e) { __check(__line(e.message), "original"); }
