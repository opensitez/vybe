// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_non_enumerable_prototype_properties
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
Object.defineProperty(proto, "hidden", { value: 100, enumerable: false });
const obj = Object.create(proto);
__check(__line(obj.hidden + "|" + ("hidden" in obj)), "100|true");
