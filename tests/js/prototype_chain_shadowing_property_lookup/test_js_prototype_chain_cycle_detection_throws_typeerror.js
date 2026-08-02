// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_cycle_detection_throws_typeerror
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

const a = {};
const b = Object.create(a);
try {
    Object.setPrototypeOf(a, b);
} catch (e) {
    __check(__line("Prototype Cycle TypeError"), "Prototype Cycle TypeError");
}
