// vybe-test: js/object_define_property_get_set_descriptors/test_js_object_define_property_symbol_key
// origin: languages/js/tests/js/test_js_object_define_property_get_set_descriptors.rs

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

const sym = Symbol("privateKey");
const obj = {};
Object.defineProperty(obj, sym, {
    value: "Secret",
    writable: true,
    enumerable: true
});
__check(__line(obj[sym]), "Secret");
__check(__line(Object.getOwnPropertySymbols(obj).length), "1");
