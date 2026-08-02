// vybe-test: js/object_descriptors/define_property_non_writable_prevents_assignment
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
Object.defineProperty(obj, "fixed", { value: 42, writable: false, configurable: true });
obj.fixed = 99;  // silently ignored (non-strict)
__check(__line(obj.fixed), "42");
__check(__line(obj.fixed === 42), "true");
