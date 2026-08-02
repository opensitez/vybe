// vybe-test: js/abstract_equality_type_coercion_rules/test_js_abstract_equality_valueof_falls_back_to_tostring
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
    valueOf() { return {}; },
    toString() { return "7"; },
};
__check(__line(obj == 7), "true");
__check(__line(obj == "7"), "true");
