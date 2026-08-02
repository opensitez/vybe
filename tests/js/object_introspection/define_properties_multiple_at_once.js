// vybe-test: js/object_introspection/define_properties_multiple_at_once
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
Object.defineProperties(obj, {
    x: { value: 10, enumerable: true, writable: true, configurable: true },
    y: { value: 20, enumerable: true, writable: true, configurable: true },
    z: { value: 30, enumerable: false, writable: false, configurable: false }
});
__check(__line(obj.x), "10");
__check(__line(obj.y), "20");
__check(__line(obj.z), "30");
__check(__line(Object.keys(obj).join(",")), "x,y");
