// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_accessor_writable_conflict_throws
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
try {
    Object.defineProperty(obj, "bad", {
        writable: true,
        set(v) {}
    });
} catch (e) {
    __check(__line("Conflict: " + e.name), "Conflict: TypeError");
}
