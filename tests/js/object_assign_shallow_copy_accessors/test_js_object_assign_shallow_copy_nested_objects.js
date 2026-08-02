// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_shallow_copy_nested_objects
// origin: languages/js/tests/js/test_js_object_assign_shallow_copy_accessors.rs

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

const nested = { count: 1 };
const source = { inner: nested };
const target = Object.assign({}, source);
target.inner.count = 99;
__check(__line(source.inner.count), "99"); // Modifying target's nested object mutates source's nested object!
