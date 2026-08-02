// vybe-test: js/error_types/type_error_basic
// origin: languages/js/tests/js/test_error_types.rs

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
    null.foo;
} catch (e) {
    __check(__line(e instanceof TypeError), "true");
    __check(__line(e.message), "Cannot read properties of null");
}
