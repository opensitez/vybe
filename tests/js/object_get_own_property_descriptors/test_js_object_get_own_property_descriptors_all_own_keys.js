// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptors_all_own_keys
// origin: languages/js/tests/js/test_js_object_get_own_property_descriptors.rs

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

const obj = { x: 1, y: 2 };
const descs = Object.getOwnPropertyDescriptors(obj);
__check(__line(Object.keys(descs).join(",")), "x,y");
__check(__line(descs.x.value + "|" + descs.y.value), "1|2");
