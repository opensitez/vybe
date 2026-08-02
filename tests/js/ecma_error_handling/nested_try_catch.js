// vybe-test: js/ecma_error_handling/nested_try_catch
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

try {
    try {
        throw new Error("inner");
    } catch (e) {
        __check(__line("inner: " + e.message), "inner: inner");
        throw new Error("rethrown");
    }
} catch (e) {
    __check(__line("outer: " + e.message), "outer: rethrown");
}
