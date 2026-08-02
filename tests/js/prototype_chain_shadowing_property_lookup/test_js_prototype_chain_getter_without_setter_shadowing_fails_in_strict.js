// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_getter_without_setter_shadowing_fails_in_strict
// origin: languages/js/tests/js/test_js_prototype_chain_shadowing_property_lookup.rs

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
    get readOnlyProp() { return "ReadOnly"; }
};
const obj = Object.create(proto);
try {
    "use strict";
    obj.readOnlyProp = "NewVal";
} catch (e) {
    __check(__line("Assign ReadOnly Prototype Getter TypeError"), "Assign ReadOnly Prototype Getter TypeError");
}
__check(__line(obj.readOnlyProp), "ReadOnly");
