// vybe-test: js/ecma_error_handling/return_from_catch_still_runs_finally
// origin: languages/js/tests/js/test_ecma_error_handling.rs

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

function test() {
    try {
        throw new Error("x");
    } catch (e) {
        __check(__line("catch"), "catch");
        return "done";
    } finally {
        __check(__line("finally"), "finally");
    }
}
__check(__line(test()), "done");
