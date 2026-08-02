// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_set_prototype_to_null_breaks_instanceof_object
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
const obj = Object.create(proto);
__check(__line((obj instanceof Object) + "|" + ("x" in obj)), "false|false");
Object.setPrototypeOf(proto, null);
    __check(__line((obj instanceof Object) + "|" + (Object.getPrototypeOf(obj) === proto)), "false|true");
