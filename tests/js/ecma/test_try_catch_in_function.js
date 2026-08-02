// vybe-test: js/ecma/test_try_catch_in_function
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

function safeDivide(a, b) {
            try {
                if (b === 0) throw "division by zero";
                return a / b;
            } catch (e) {
                return e;
            }
        }
        __check(__line(safeDivide(10, 2)), "5");
        __check(__line(safeDivide(10, 0)), "division by zero");
