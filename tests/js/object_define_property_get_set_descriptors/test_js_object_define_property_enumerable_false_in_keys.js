// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_enumerable_false_in_keys
// origin: languages/js/tests/js/test_js_object_define_property_get_set_descriptors.rs

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

const obj = { a: 1 };
Object.defineProperty(obj, "hidden", {
    value: 2,
    enumerable: false
});
__check(__line(Object.keys(obj).join(",")), "a");
__check(__line(obj.hidden), "2");
