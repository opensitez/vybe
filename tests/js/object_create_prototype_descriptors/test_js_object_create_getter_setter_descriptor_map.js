// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_getter_setter_descriptor_map
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

let store = 0;
const obj = Object.create(null, {
    val: {
        get() { return store; },
        set(v) { store = v * 2; },
        enumerable: true
    }
});
obj.val = 5;
__check(__line(obj.val), "10");
