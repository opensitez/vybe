// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_getter_setter_accessors
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

const obj = { _val: 0 };
Object.defineProperty(obj, "val", {
    get() { return this._val * 2; },
    set(v) { this._val = v + 5; },
    enumerable: true,
    configurable: true
});
obj.val = 10;
__check(__line(obj._val), "15");
__check(__line(obj.val), "30");
