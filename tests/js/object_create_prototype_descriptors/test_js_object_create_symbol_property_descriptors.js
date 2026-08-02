// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_symbol_property_descriptors
// origin: languages/js/tests/js/test_js_object_create_prototype_descriptors.rs

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

const sym = Symbol("symKey");
const obj = Object.create(null, {
    [sym]: { value: "symbolVal", enumerable: true }
});
__check(__line(obj[sym] + "|" + Object.getOwnPropertySymbols(obj).length), "symbolVal|1");
