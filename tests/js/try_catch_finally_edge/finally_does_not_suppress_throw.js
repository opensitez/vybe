// vybe-test: js/try_catch_finally_edge/finally_does_not_suppress_throw
// origin: languages/js/tests/js/test_try_catch_finally_edge.rs

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

let caught = null;
try {
    try {
        throw new Error("original");
    } finally {
        // finally without catch — error still propagates
        __check(__line("finally runs"), "finally runs");
    }
} catch (e) {
    caught = e.message;
}
__check(__line(caught), "original");
