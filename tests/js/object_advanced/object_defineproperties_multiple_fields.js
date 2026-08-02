// vybe-test: js/object_advanced/object_defineproperties_multiple_fields
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
Object.defineProperties(obj, {
    a: { value: 1, enumerable: true },
    b: { value: 2, enumerable: true }
});
__check(__line(obj.a), "1");
__check(__line(obj.b), "2");
__check(__line(Object.keys(obj).join(",")), "a,b");
