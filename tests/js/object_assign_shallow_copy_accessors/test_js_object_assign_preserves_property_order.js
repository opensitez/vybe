// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_preserves_property_order
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

const source = { 2: "b", 1: "a", c: "c" };
const target = Object.assign({}, source);
__check(__line(Object.keys(target).join(",")), "1,2,c");
