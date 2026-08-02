// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_instanceof_custom_symbol_has_instance
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

const OddChecker = {
    [Symbol.hasInstance](val) {
        return typeof val === "number" && val % 2 !== 0;
    }
};
__check(__line((5 instanceof OddChecker) + "|" + (4 instanceof OddChecker)), "true|false");
