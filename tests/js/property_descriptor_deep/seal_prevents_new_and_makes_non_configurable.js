// vybe-test: js/property_descriptor_deep/seal_prevents_new_and_makes_non_configurable
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

const obj = { x: 1 };
Object.seal(obj);
obj.y = 2; // silently fails
__check(__line(obj.y), "undefined");
obj.x = 99; // still writable
__check(__line(obj.x), "99");
const keyCount = Object.keys(obj).length;
obj.z = 3; // try adding another property
__check(__line(Object.keys(obj).length === keyCount), "true"); // sealed — no new keys
