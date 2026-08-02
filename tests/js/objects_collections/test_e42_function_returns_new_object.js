// vybe-test: js/objects_collections/test_e42_function_returns_new_object
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

function makePoint(x, y) { return { x: x, y: y }; }
        let p = makePoint(3, 4);
        __check(__line(p.x, p.y), "3 4");
