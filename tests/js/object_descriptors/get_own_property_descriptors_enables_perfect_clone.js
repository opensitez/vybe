// vybe-test: js/object_descriptors/get_own_property_descriptors_enables_perfect_clone
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

const orig = {};
Object.defineProperty(orig, "x", { value: 42, writable: false, enumerable: true, configurable: false });
const clone = {};
Object.defineProperty(clone, "x", { value: orig.x, writable: false, enumerable: true, configurable: false });
__check(__line(clone.x), "42");
clone.x = 99;  // silently ignored (writable: false)
__check(__line(clone.x), "42");
