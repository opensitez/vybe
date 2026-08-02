// vybe-test: js/ecma_error_handling/throw_plain_object_and_read_property
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
    throw { code: 500, message: "server" };
} catch (e) {
    __check(__line(e.code), "500");
    __check(__line(e.message), "server");
}
