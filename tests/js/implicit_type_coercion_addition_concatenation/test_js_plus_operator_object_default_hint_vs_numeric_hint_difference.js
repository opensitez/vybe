// vybe-test: js/implicit_type_coercion_addition_concatenation/test_js_plus_operator_object_default_hint_vs_numeric_hint_difference
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

const obj = { valueOf: () => "10" };
__check(__line(`${obj + 1}:${obj - 1}`), "101:9");
