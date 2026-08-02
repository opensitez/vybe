// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_getter_setter_descriptors_defined_via_define_property
// origin: languages/js/tests/js/test_js_property_accessors_getters_setters_inheritance.rs

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
    get() { return this._val; },
    set(v) { this._val = v * 10; },
    enumerable: true,
    configurable: true
});
obj.val = 5;
__check(__line(obj.val), "50");
