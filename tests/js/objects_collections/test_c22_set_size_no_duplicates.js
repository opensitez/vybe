// vybe-test: js/objects_collections/test_c22_set_size_no_duplicates
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
        s.add("a");
        s.add("b");
        s.add("a");
        s.add("c");
        s.add("b");
        __check(__line(s.size), "3");
