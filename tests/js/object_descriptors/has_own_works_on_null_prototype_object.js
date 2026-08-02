// vybe-test: js/object_descriptors/has_own_works_on_null_prototype_object
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

const obj = Object.create(null);
obj.x = 42;
__check(__line(Object.hasOwn(obj, "x")), "true");
__check(__line(Object.hasOwn(obj, "y")), "false");
