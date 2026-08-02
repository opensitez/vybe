// vybe-test: js/object_has_own_vs_has_own_property/test_js_has_own_property_fails_on_null_prototype_objects
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

const nullProtoObj = Object.create(null);
nullProtoObj.key = 100;
try {
    nullProtoObj.hasOwnProperty("key");
} catch (e) {
    __check(__line("hasOwnProperty Null Prototype TypeError"), "hasOwnProperty Null Prototype TypeError");
}
