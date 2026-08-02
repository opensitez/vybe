// vybe-test: js/type_coercion_deep/abstract_eq_boolean_coerces_to_number
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

__check(__line(true == 1), "true");
__check(__line(false == 0), "true");
__check(__line(true == "1"), "true");
__check(__line(false == ""), "true");
