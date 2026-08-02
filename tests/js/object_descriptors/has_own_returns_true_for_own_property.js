// vybe-test: js/object_descriptors/has_own_returns_true_for_own_property
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

const obj = { a: 1 };
__check(__line(Object.hasOwn(obj, "a")), "true");
__check(__line(Object.hasOwn(obj, "toString")), "false");
