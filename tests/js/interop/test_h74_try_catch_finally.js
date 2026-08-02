// vybe-test: js/interop/test_h74_try_catch_finally
// origin: languages/js/tests/js/js_interop_test.rs

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
            log += "try ";
            throw "oops";
        } catch (e) {
            log += "catch(" + e + ") ";
        } finally {
            log += "finally";
        }
        __check(__line(log), "try catch(oops) finally");

        let log2 = "";
        try {
            log2 += "try ";
        } catch (e) {
            log2 += "catch ";
        } finally {
            log2 += "finally";
        }
        __check(__line(log2), "try finally");
