// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_property_descriptor_details
// origin: languages/js/tests/js/test_js_object_has_own_vs_has_own_property.rs

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

const desc = Object.getOwnPropertyDescriptor(Object, "hasOwn");
__check(__line(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Object.hasOwn.length}`), "true:false:true:2");
