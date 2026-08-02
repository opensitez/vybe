// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_class_static_and_instance_fields
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

class Widget {
    static staticProp = 1;
    instanceProp = 2;
}
const w = new Widget();
__check(__line(`${Object.hasOwn(Widget, "staticProp")}:${Object.hasOwn(w, "instanceProp")}:${Object.hasOwn(w, "staticProp")}`), "true:true:false");
