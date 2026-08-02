// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_static_private_brand_check
// origin: languages/js/tests/js/test_js_class_private_in_operator_brand_check.rs

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

class System {
    static #config = {};
    static isSystemClass(target) {
        return #config in target;
    }
}
__check(__line(System.isSystemClass(System) + "|" + System.isSystemClass(new System())), "true|false");
