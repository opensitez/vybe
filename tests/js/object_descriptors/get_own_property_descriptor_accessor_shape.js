// vybe-test: js/object_descriptors/get_own_property_descriptor_accessor_shape
// origin: languages/js/tests/js/test_object_descriptors.rs

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
Object.defineProperty(obj, "n", {
    get() { return 7; },
    set(v) { this._n = v; },
    enumerable: false,
    configurable: true
});
const d = Object.getOwnPropertyDescriptor(obj, "n");
__check(__line(typeof d.get), "function");
__check(__line(typeof d.set), "function");
__check(__line(d.enumerable), "false");
__check(__line(d.configurable), "true");
