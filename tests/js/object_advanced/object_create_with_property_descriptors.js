// vybe-test: js/object_advanced/object_create_with_property_descriptors
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

let obj = Object.create({}, {
    x: { value: 5, enumerable: true },
    y: { value: 7, enumerable: true }
});
__check(__line(obj.x + obj.y), "12");
__check(__line(Object.keys(obj).join(",")), "x,y");
