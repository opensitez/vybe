// vybe-test: js/objects_collections/test_c23_set_delete
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

let s = new Set();
        s.add(10);
        s.add(20);
        s.delete(10);
        __check(__line(s.has(10), s.size), "false 1");
