// vybe-test: js/implicit_type_coercion_addition_concatenation/test_js_plus_operator_string_and_number
// origin: languages/js/tests/js/test_js_implicit_type_coercion_addition_concatenation.rs

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

__check(__line([
    "5" + 5,
    5 + "5",
    "5" + 5 + 5,
    5 + 5 + "5"
].join("|")), "55|55|555|105");
