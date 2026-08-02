// vybe-test: js/object_methods_deep/object_create_with_property_descriptors
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const obj = Object.create({}, {
    x: { value: 10, writable: true, enumerable: true, configurable: true },
    y: { value: 20, writable: false, enumerable: true, configurable: true }
});
__check(__line(obj.x), "10");
__check(__line(obj.y), "20");
