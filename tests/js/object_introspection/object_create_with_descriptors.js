// vybe-test: js/object_introspection/object_create_with_descriptors
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

const obj = Object.create(Object.prototype, {
    name: { value: "Alice", enumerable: true, writable: true, configurable: true },
    age:  { value: 30,      enumerable: true, writable: false, configurable: false }
});
__check(__line(obj.name), "Alice");
__check(__line(obj.age), "30");
obj.age = 99;
__check(__line(obj.age), "30");
