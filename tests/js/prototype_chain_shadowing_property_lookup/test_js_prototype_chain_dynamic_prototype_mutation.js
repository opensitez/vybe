// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_dynamic_prototype_mutation
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

const proto1 = { v: "P1" };
const proto2 = { v: "P2" };
const obj = Object.create(proto1);
__check(__line(obj.v), "P1");
Object.setPrototypeOf(obj, proto2);
__check(__line(obj.v), "P2");
