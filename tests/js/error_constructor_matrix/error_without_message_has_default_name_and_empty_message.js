// vybe-test: js/error_constructor_matrix/error_without_message_has_default_name_and_empty_message
// origin: languages/js/tests/js/test_error_constructor_matrix.rs

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

const e = new Error();
__check(__line(e.name), "Error");
__check(__line(e.message), "");
