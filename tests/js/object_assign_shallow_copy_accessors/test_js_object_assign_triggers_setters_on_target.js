// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_triggers_setters_on_target
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

let targetSetterCalled = false;
const target = {
    set a(v) { targetSetterCalled = true; this._a = v * 2; }
};
Object.assign(target, { a: 10 });
__check(__line(targetSetterCalled + "|" + target._a), "true|20");
