// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_property_descriptor_defaults_on_target
// origin: languages/js/tests/js/test_js_object_assign_shallow_copy_accessors.rs

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

const target = Object.assign({}, { key: "val" });
const desc = Object.getOwnPropertyDescriptor(target, "key");
__check(__line(`${desc.writable}:${desc.enumerable}:${desc.configurable}`), "true:true:true");
