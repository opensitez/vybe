// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_evaluates_getters_during_copy
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

const source = {
    get val() { return "GetterVal"; }
};
const target = Object.assign({}, source);
const desc = Object.getOwnPropertyDescriptor(target, "val");
__check(__line(target.val + "|hasGetter=" + (desc.get !== undefined)), "GetterVal|hasGetter=false");
