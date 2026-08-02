// vybe-test: js/ecma/test_default_param_arrow
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

let add = (a, b = 1) => a + b;
        __check(__line(add(5)), "6");
        __check(__line(add(5, 10)), "15");
