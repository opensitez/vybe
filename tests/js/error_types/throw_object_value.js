// vybe-test: js/error_types/throw_object_value
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
    throw { code: 500, msg: "internal error" };
} catch (e) {
    __check(__line(e.code), "500");
    __check(__line(e.msg), "internal error");
}
