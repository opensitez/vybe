// vybe-test: js/primitive_conversion_builtins/number_octal_string_parses
// origin: languages/js/tests/js/test_primitive_conversion_builtins.rs

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

__check(__line(Number("0o10")), "8");
