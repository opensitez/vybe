// vybe-test: js/object_introspection/get_own_property_descriptors_returns_all
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

const obj = { a: 1, b: 2 };
Object.defineProperty(obj, "c", { value: 3, enumerable: false, writable: true, configurable: true });
const descs = Object.getOwnPropertyDescriptors(obj);
__check(__line(descs.a.value), "1");
__check(__line(descs.b.value), "2");
__check(__line(descs.c.value), "3");
__check(__line(descs.c.enumerable), "false");
