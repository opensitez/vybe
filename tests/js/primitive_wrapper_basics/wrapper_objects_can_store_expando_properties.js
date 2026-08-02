// vybe-test: js/primitive_wrapper_basics/wrapper_objects_can_store_expando_properties
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

const s = new String("hi");
s.extra = 1;
const n = new Number(2);
n.tag = "x";
__check(__line(s.extra), "1");
__check(__line(n.tag), "x");
