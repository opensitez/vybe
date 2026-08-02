// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_default_descriptor_boolean_flags_are_false
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
    prop: { value: "defaultFlags" }
});
const desc = Object.getOwnPropertyDescriptor(obj, "prop");
__check(__line(`${desc.writable}:${desc.enumerable}:${desc.configurable}`), "false:false:false");
