// vybe-test: js/try_catch_finally_edge/rethrow_propagates_original
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

function inner() { throw new TypeError("inner"); }
function outer() {
    try { inner(); }
    catch (e) {
        if (!(e instanceof TypeError)) throw e;
        __check(__line("handled: " + e.message), "handled: inner");
    }
}
outer();
