// vybe-test: js/objects_collections/test_a09_property_shorthand
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

let x = 100;
        let y = 200;
        let obj = { x, y };
        __check(__line(obj.x, obj.y), "100 200");
