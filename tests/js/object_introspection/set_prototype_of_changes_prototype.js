// vybe-test: js/object_introspection/set_prototype_of_changes_prototype
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

const newProto = { greet() { return "hello from newProto"; } };
const obj = { x: 1 };
Object.setPrototypeOf(obj, newProto);
__check(__line(Object.getPrototypeOf(obj) === newProto), "true");
__check(__line(obj.greet()), "hello from newProto");
