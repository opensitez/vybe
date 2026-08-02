// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_overriding_inherited_getter_with_own_data_property
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
    get name() { return "ProtoName"; }
};
const obj = Object.create(proto);
obj.name = "OwnName"; // Shadowing getter with own property fails in non-strict mode if non-writable in proto!
__check(__line(obj.name), "ProtoName");
