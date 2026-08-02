// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_inherited_getter_evaluates_with_receiver_this
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
    get value() { return this._val * 2; }
};
const obj = Object.create(proto);
obj._val = 10;
__check(__line(obj.value), "20");
