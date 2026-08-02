// vybe-test: js/error_handling_advanced/error_in_finally
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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
        throw new Error("original");
    } finally {
        return "from finally";
    }
}
__check(__line(test()), "from finally");
