// vybe-test: js/object_methods_deep/get_own_property_descriptor_accessor
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const obj = {};
let _val = 0;
Object.defineProperty(obj, "v", {
    get() { return _val; },
    set(x) { _val = x; },
    enumerable: true,
    configurable: true
});
const desc = Object.getOwnPropertyDescriptor(obj, "v");
__check(__line(typeof desc.get), "function");
__check(__line(typeof desc.set), "function");
__check(__line("value" in desc), "false");
