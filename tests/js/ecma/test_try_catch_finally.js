// vybe-test: js/ecma/test_try_catch_finally
// origin: languages/js/tests/js/js_ecma_test.rs

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

let log = "";
        try {
            log = log + "try ";
            throw "err";
        } catch (e) {
            log = log + "catch ";
        } finally {
            log = log + "finally";
        }
        __check(__line(log), "try catch finally");
