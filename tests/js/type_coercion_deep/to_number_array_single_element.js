// vybe-test: js/type_coercion_deep/to_number_array_single_element
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

__check(__line(+[]), "0");
__check(__line(+[42]), "42");
__check(__line(isNaN(+[1,2])), "true");
