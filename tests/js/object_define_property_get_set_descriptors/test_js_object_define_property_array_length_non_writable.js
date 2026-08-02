// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_array_length_non_writable
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

const arr = [10, 20];
Object.defineProperty(arr, "length", { writable: false });
try {
    arr.push(30);
} catch (e) {
    __check(__line("Length Non-Writable Error"), "Length Non-Writable Error");
}
