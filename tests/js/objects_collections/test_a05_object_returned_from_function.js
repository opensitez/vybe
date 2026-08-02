// vybe-test: js/objects_collections/test_a05_object_returned_from_function
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

function make() { return { x: 10, y: 20 }; }
        let r = make();
        __check(__line(r.x, r.y), "10 20");
