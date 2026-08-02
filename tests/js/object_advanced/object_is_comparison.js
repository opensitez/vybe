// vybe-test: js/object_advanced/object_is_comparison
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

__check(__line(Object.is(42, 42)), "true");
__check(__line(Object.is("foo", "foo")), "true");
__check(__line(Object.is(NaN, NaN)), "true");
__check(__line(Object.is(0, -0)), "false");
__check(__line(Object.is(null, undefined)), "false");
