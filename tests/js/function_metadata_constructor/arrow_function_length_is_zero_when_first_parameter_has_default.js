// vybe-test: js/function_metadata_constructor/arrow_function_length_is_zero_when_first_parameter_has_default
// origin: languages/js/tests/js/test_function_metadata_constructor.rs

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

const fn = (a = 1, b) => a + b;
__check(__line(fn.length), "0");
