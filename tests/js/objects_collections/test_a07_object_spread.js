// vybe-test: js/objects_collections/test_a07_object_spread
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

let base = { a: 1, b: 2 };
        let extended = { ...base, c: 3 };
        __check(__line(extended.a, extended.b, extended.c), "1 2 3");
