// vybe-test: js/error_types_deep/range_error_from_invalid_array_length
// origin: languages/js/tests/js/test_error_types_deep.rs

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

try { new Array(-1); } catch (e) {
    __check(__line(e instanceof RangeError), "true");
    __check(__line(e.name), "RangeError");
}
