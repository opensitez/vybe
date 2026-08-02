// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_non_writable_prototype_property_shadowing_blocked
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

const proto = {};
Object.defineProperty(proto, "fixed", { value: 10, writable: false });
const obj = Object.create(proto);

try {
    "use strict";
    obj.fixed = 20; // Cannot shadow non-writable property on prototype via normal assignment!
} catch (e) {
    __check(__line("Shadowing Non-Writable Prototype Property TypeError"), "Shadowing Non-Writable Prototype Property TypeError");
}
__check(__line(obj.fixed), "10");
