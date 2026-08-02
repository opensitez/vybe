// vybe-test: js/objects_collections/test_a02_object_dynamic_property_add
// origin: languages/js/tests/js/js_objects_collections_test.rs

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

let obj = { x: 1 };
        obj.y = 2;
        obj.z = 3;
        __check(__line(obj.x, obj.y, obj.z), "1 2 3");
