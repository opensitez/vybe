// vybe-test: js/object_introspection/get_prototype_of_returns_prototype
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

const proto = { kind: "proto" };
const obj = Object.create(proto);
__check(__line(Object.getPrototypeOf(obj) === proto), "true");
__check(__line(obj.kind), "proto");
