// vybe-test: js/numeric_coercion_deep/number_to_string_radix
// origin: languages/js/tests/js/test_numeric_coercion_deep.rs

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

__check(__line((255).toString(16)), "ff");
__check(__line((10).toString(2)), "1010");
__check(__line((31).toString(8)), "37");
