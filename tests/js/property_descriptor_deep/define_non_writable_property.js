// vybe-test: js/property_descriptor_deep/define_non_writable_property
// origin: languages/js/tests/js/test_property_descriptor_deep.rs

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
Object.defineProperty(obj, "x", { value: 42, writable: false, configurable: true, enumerable: true });
obj.x = 99; // silently fails in sloppy mode
__check(__line(obj.x), "42");
