// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_null_prototype_object_requires_custom_methods
// origin: languages/js/tests/js/test_js_object_to_primitive_valueof_tostring_order.rs

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

const obj = Object.create(null);
try {
    Number(obj);
} catch (e) {
    __check(__line("Null Prototype Object ToPrimitive TypeError"), "Null Prototype Object ToPrimitive TypeError");
}
