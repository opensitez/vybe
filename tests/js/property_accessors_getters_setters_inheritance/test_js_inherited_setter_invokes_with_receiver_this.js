// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_inherited_setter_invokes_with_receiver_this
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

const proto = {
    set value(v) { this._val = v + 1; }
};
const obj = Object.create(proto);
obj.value = 5;
__check(__line(obj._val + "|hasOwn=" + Object.hasOwn(obj, "_val")), "6|hasOwn=true");
