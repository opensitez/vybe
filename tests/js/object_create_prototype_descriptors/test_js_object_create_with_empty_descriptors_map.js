// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_with_empty_descriptors_map
// origin: languages/js/tests/js/test_js_object_create_prototype_descriptors.rs

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

const obj = Object.create({}, {});
__check(__line(Object.keys(obj).length), "0");
