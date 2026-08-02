// vybe-test: js/object_descriptors/get_own_property_descriptor_data_descriptor
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

const obj = { x: 42 };
const d = Object.getOwnPropertyDescriptor(obj, "x");
__check(__line(d.value), "42");
__check(__line(d.writable), "true");
__check(__line(d.enumerable), "true");
__check(__line(d.configurable), "true");
