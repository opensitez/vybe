// vybe-test: js/class_extends_super_constructor_call/test_js_class_extends_expression_dynamic_inheritance
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

function createBaseClass(prefix) {
    return class {
        getPrefix() { return prefix; }
    };
}
class DynamicChild extends createBaseClass("[DYN]") {}
__check(__line(new DynamicChild().getPrefix()), "[DYN]");
