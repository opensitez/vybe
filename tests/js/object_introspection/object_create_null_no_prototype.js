// vybe-test: js/object_introspection/object_create_null_no_prototype
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

const obj = Object.create(null);
obj.foo = "bar";
__check(__line(obj.foo), "bar");
__check(__line(Object.getPrototypeOf(obj)), "null");
