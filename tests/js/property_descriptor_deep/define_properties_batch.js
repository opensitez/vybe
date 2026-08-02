// vybe-test: js/property_descriptor_deep/define_properties_batch
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
Object.defineProperties(obj, {
    a: { value: 1, enumerable: true, configurable: true, writable: true },
    b: { value: 2, enumerable: true, configurable: true, writable: true },
});
__check(__line(obj.a + obj.b), "3");
