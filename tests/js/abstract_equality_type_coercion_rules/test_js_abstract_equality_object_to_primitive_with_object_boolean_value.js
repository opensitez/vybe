// vybe-test: js/abstract_equality_type_coercion_rules/test_js_abstract_equality_object_to_primitive_with_object_boolean_value
// origin: languages/js/tests/js/test_js_abstract_equality_type_coercion_rules.rs

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

const obj = { valueOf: () => false };
__check(__line(`${obj == false}:${obj == 0}`), "true:true");
