// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptors_shallow_copy
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

const orig = { a: 1 };
const descs = Object.getOwnPropertyDescriptors(orig);
descs.a.value = 99;
__check(__line(orig.a), "1"); // Original target property value unaltered
