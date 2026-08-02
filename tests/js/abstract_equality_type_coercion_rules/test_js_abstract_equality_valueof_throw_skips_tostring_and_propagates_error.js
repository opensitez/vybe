// vybe-test: js/abstract_equality_type_coercion_rules/test_js_abstract_equality_valueof_throw_skips_tostring_and_propagates_error
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

const obj = {
    valueOf() {
        throw new Error("valueOfBoom");
    },
    toString() {
        return "7";
    }
};
try {
    console.log(obj == 7);
} catch (e) {
    console.log("throws");
}
