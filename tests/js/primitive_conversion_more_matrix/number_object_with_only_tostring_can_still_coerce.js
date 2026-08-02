// vybe-test: js/primitive_conversion_more_matrix/number_object_with_only_tostring_can_still_coerce
// origin: languages/js/tests/js/test_primitive_conversion_more_matrix.rs

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

__check(__line(Number({ toString() { return " 9 "; } })), "9");
