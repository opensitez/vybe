// vybe-test: js/objects_collections/test_a01_object_literal_with_properties
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

let obj = { name: "Alice", age: 30 };
        __check(__line(obj.name, obj.age), "Alice 30");
