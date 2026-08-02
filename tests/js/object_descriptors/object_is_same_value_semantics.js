// vybe-test: js/object_descriptors/object_is_same_value_semantics
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

__check(__line(Object.is(1, 1)), "true");
__check(__line(Object.is("a", "a")), "true");
__check(__line(Object.is(null, null)), "true");
__check(__line(Object.is(undefined, undefined)), "true");
