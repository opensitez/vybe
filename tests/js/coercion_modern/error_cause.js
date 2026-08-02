// vybe-test: js/coercion_modern/error_cause
// origin: languages/js/tests/js/test_coercion_modern.rs

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
        throw new Error("original");
    } catch (e) {
        throw new Error("wrapped", { cause: e });
    }
} catch (e) {
    __check(__line(e.message), "wrapped");
    __check(__line(e.cause.message), "original");
}
