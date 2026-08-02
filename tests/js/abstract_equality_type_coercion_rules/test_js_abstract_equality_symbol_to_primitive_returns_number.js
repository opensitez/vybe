// vybe-test: js/abstract_equality_type_coercion_rules/test_js_abstract_equality_symbol_to_primitive_returns_number
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

const value = {
    [Symbol.toPrimitive]() {
        return 1;
    }
};
__check(__line(`${value == 1}:${value == "1"}`), "true:true");
