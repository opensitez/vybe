// vybe-test: js/intl_ecma402_surface/test_js_intl_primitive_integration

function assert(cond, msg) {
    if (!cond) {
        throw new Error(msg);
    }
}

assert((1234567.89).toLocaleString("en-US").includes(","), "number toLocaleString groups");
assert("ä".localeCompare("z", "de") < 0, "localeCompare locale argument");

let symbolString = false;
try {
    "a".localeCompare(Symbol("a"));
} catch (e) {
    symbolString = true;
}
assert(symbolString, "localeCompare Symbol throws");
console.log("ok");
