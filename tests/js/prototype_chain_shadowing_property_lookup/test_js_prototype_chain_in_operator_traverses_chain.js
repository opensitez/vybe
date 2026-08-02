// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_in_operator_traverses_chain
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

const proto = { key: 1 };
const obj = Object.create(proto);
__check(__line(("key" in obj) + "|" + Object.hasOwn(obj, "key")), "true|false");
