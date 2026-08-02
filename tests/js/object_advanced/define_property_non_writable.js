// vybe-test: js/object_advanced/define_property_non_writable
// origin: languages/js/tests/js/test_object_advanced.rs

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

let obj = {};
Object.defineProperty(obj, "PI", {
    value: 3.14159,
    writable: false,
    enumerable: true
});
__check(__line(obj.PI), "3.14159");
obj.PI = 0;
__check(__line(obj.PI), "3.14159");
