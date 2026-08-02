// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_define_property_bypasses_non_writable_prototype
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

// Object.defineProperty directly defines own property, bypassing prototype check!
Object.defineProperty(obj, "fixed", { value: 20, writable: true });
__check(__line(obj.fixed + "|" + proto.fixed), "20|10");
