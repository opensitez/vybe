// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_properties_multiple
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

const obj = {};
Object.defineProperties(obj, {
    x: { value: 10, writable: true, enumerable: true },
    y: { get() { return this.x * 2; }, enumerable: true }
});
__check(__line(obj.x + "," + obj.y), "10,20");
obj.x = 20;
__check(__line(obj.y), "40");
