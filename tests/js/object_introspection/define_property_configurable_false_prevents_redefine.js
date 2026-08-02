// vybe-test: js/object_introspection/define_property_configurable_false_prevents_redefine
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
Object.defineProperty(obj, "locked", { value: 1, configurable: false, writable: false, enumerable: true });
let threw = false;
try {
    Object.defineProperty(obj, "locked", { value: 2 });
} catch (e) {
    threw = true;
}
__check(__line(threw), "true");
__check(__line(obj.locked), "1");
