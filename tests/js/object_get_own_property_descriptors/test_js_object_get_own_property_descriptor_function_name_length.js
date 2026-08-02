// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptor_function_name_length
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

function testFn(a, b) {}
const descName = Object.getOwnPropertyDescriptor(testFn, "name");
const descLen = Object.getOwnPropertyDescriptor(testFn, "length");
__check(__line(descName.value + "|" + descLen.value + "|" + descName.writable), "testFn|2|false");
