// vybe-test: js/ecma/test_try_catch_nested
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

let result = "";
        try {
            try {
                throw "inner";
            } catch (e) {
                result = result + e + " ";
                throw "outer";
            }
        } catch (e) {
            result = result + e;
        }
        __check(__line(result), "inner outer");
