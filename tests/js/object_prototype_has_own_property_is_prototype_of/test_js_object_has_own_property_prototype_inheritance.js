// vybe-test: js/object_prototype_has_own_property_is_prototype_of/test_js_object_has_own_property_prototype_inheritance
// origin: languages/js/tests/js/test_js_object_prototype_has_own_property_is_prototype_of.rs

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

const parent = { parentKey: "parentVal" };
const child = Object.create(parent);
child.childKey = "childVal";

__check(__line(child.hasOwnProperty("childKey")), "true");
__check(__line(child.hasOwnProperty("parentKey")), "false");
