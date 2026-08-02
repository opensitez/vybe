// vybe-test: js/property_descriptor_deep/get_own_property_descriptor_returns_all_attributes
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
const desc = Object.getOwnPropertyDescriptor(obj, "x");
__check(__line(desc.value), "1");
__check(__line(desc.writable), "true");
__check(__line(desc.enumerable), "true");
__check(__line(desc.configurable), "true");
