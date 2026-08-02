// vybe-test: js/object_immutability/seal_makes_existing_properties_non_configurable
// origin: languages/js/tests/js/test_object_immutability.rs

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

const obj = Object.seal({ a: 1, b: 2 });
const d = Object.getOwnPropertyDescriptor(obj, "a");
__check(__line(d.configurable), "false");
__check(__line(d.writable), "true");
