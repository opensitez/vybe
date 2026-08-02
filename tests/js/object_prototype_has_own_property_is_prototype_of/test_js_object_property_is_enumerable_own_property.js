// vybe-test: js/object_prototype_has_own_property_is_prototype_of/test_js_object_property_is_enumerable_own_property
// origin: languages/js/tests/js/test_js_object_prototype_has_own_property_is_prototype_of.rs

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

const obj = { visible: 1 };
Object.defineProperty(obj, "invisible", { value: 2, enumerable: false });

__check(__line(obj.propertyIsEnumerable("visible")), "true");
__check(__line(obj.propertyIsEnumerable("invisible")), "false");
