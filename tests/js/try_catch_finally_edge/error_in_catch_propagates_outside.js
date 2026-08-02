// vybe-test: js/try_catch_finally_edge/error_in_catch_propagates_outside
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

let caught = false;
try {
    try {
        throw new Error("first");
    } catch {
        throw new Error("second"); // error in catch
    }
} catch (e) {
    caught = true;
    __check(__line(e.message), "second");
}
__check(__line(caught), "true");
