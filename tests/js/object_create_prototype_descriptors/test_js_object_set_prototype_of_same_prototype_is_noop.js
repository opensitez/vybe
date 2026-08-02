// vybe-test: js/object_create_prototype_descriptors/test_js_object_set_prototype_of_same_prototype_is_noop
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

const proto = { x: 10 };
const obj = Object.create(proto);
const res = Object.setPrototypeOf(obj, proto);
__check(__line(res === obj), "true");
