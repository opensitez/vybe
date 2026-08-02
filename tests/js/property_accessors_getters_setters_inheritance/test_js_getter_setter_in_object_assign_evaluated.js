// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_getter_setter_in_object_assign_evaluated
// origin: languages/js/tests/js/test_js_property_accessors_getters_setters_inheritance.rs

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

const source = {
    get val() { return "EvaluatedValue"; }
};
const target = Object.assign({}, source);
const desc = Object.getOwnPropertyDescriptor(target, "val");
__check(__line(target.val + "|isDataProperty=" + (desc.get === undefined)), "EvaluatedValue|isDataProperty=true");
