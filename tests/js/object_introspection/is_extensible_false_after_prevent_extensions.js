// vybe-test: js/object_introspection/is_extensible_false_after_prevent_extensions
// origin: languages/js/tests/js/test_object_introspection.rs

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
__check(__line(Object.isExtensible(obj)), "true");
Object.preventExtensions(obj);
__check(__line(Object.isExtensible(obj)), "false");
