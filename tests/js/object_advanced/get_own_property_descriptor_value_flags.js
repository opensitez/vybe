// vybe-test: js/object_advanced/get_own_property_descriptor_value_flags
// origin: languages/js/tests/js/test_object_advanced.rs

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

let obj = { a: 1 };
let d = Object.getOwnPropertyDescriptor(obj, "a");
__check(__line(d.value), "1");
__check(__line(d.writable), "true");
__check(__line(d.enumerable), "true");
__check(__line(d.configurable), "true");
