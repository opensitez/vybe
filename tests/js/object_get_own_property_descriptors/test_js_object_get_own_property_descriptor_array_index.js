// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptor_array_index
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

const arr = ["first", "second"];
const desc0 = Object.getOwnPropertyDescriptor(arr, 0);
const descLen = Object.getOwnPropertyDescriptor(arr, "length");
__check(__line(desc0.value + "|" + descLen.value + "|" + descLen.writable), "first|2|true");
