// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_array_length_truncation
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

const arr = [1, 2, 3, 4, 5];
Object.defineProperty(arr, "length", { value: 2 });
__check(__line(arr.length + "|" + arr.join(",")), "2|1,2");
