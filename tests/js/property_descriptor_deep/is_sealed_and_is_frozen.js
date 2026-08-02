// vybe-test: js/property_descriptor_deep/is_sealed_and_is_frozen
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
__check(__line(Object.isSealed(obj)), "false");  // false
__check(__line(Object.isFrozen(obj)), "false");  // false (empty non-extensible would be frozen/sealed)
Object.seal(obj);
__check(__line(Object.isSealed(obj)), "true");  // true
