// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_in_operator_works_with_string_object_wrappers
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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

__check(__line("0" in new String("abc")), "true");
__check(__line(0 in new String("abc")), "true");
__check(__line("length" in new String("")), "true");
