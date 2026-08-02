// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptor_primitive_coercion
// origin: languages/js/tests/js/test_js_object_get_own_property_descriptors.rs

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

const descNumber = Object.getOwnPropertyDescriptor(100, "toString");
const descString = Object.getOwnPropertyDescriptor("abc", "length");
__check(__line(descNumber === undefined), "true");
__check(__line(descString.value + "|" + descString.writable + "|" + descString.enumerable), "3|false|false");
