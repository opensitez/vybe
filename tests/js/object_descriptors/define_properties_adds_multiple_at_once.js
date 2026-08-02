// vybe-test: js/object_descriptors/define_properties_adds_multiple_at_once
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
Object.defineProperties(obj, {
    a: { value: 1, enumerable: true, writable: true, configurable: true },
    b: { value: 2, enumerable: true, writable: true, configurable: true },
    c: { value: 3, enumerable: false, writable: true, configurable: true }
});
const keys = Object.keys(obj);
__check(__line(keys.sort().join(",")), "a,b");
__check(__line(obj.c), "3");
