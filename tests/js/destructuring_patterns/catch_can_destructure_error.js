// vybe-test: js/destructuring_patterns/catch_can_destructure_error
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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
    throw { code: 404, message: "not found" };
} catch ({ code, message }) {
    __check(__line(code), "404");
    __check(__line(message), "not found");
}
