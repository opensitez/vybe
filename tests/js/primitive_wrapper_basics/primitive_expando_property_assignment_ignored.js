// vybe-test: js/primitive_wrapper_basics/primitive_expando_property_assignment_ignored
// origin: languages/js/tests/js/test_primitive_wrapper_basics.rs

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

const s = "hi";
s.foo = 123;
__check(__line(s.foo === undefined), "true");
