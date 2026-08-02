// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_shadowing_inherited_getter_via_define_property
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
Object.defineProperty(obj, "name", { value: "OwnName", writable: true });
__check(__line(obj.name), "OwnName");
