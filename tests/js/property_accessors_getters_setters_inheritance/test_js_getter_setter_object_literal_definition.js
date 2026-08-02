// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_getter_setter_object_literal_definition
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

const obj = {
    _age: 20,
    get age() { return this._age; },
    set age(v) { this._age = v; }
};
obj.age = 25;
__check(__line(obj.age), "25");
