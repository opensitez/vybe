// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_with_property_descriptors_map
// origin: languages/js/tests/js/test_js_object_create_prototype_descriptors.rs

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

const obj = Object.create(null, {
    x: { value: 10, writable: true, enumerable: true, configurable: true },
    y: { value: 20, writable: false, enumerable: false, configurable: false }
});
__check(__line(`${obj.x}:${obj.y}:${Object.keys(obj).join(",")}`), "10:20:x");
