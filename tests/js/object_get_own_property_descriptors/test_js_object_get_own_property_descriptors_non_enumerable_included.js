// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptors_non_enumerable_included
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

const obj = {};
Object.defineProperty(obj, "secret", { value: 42, enumerable: false });
const descs = Object.getOwnPropertyDescriptors(obj);
__check(__line(descs.secret.value + "|" + descs.secret.enumerable), "42|false");
