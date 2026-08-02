// vybe-test: js/interop/test_h75_throw_custom_error
// origin: languages/js/tests/js/js_interop_test.rs

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
            throw { message: "custom error", code: 42 };
        } catch (e) {
            __check(__line(e.message, e.code), "custom error 42");
        }
