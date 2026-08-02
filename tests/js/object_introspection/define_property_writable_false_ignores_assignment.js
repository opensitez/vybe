// vybe-test: js/object_introspection/define_property_writable_false_ignores_assignment
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

const obj = {};
Object.defineProperty(obj, "CONST", { value: 100, writable: false, enumerable: true, configurable: false });
obj.CONST = 999;
__check(__line(obj.CONST), "100");
