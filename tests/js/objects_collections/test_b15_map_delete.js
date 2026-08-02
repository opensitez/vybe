// vybe-test: js/objects_collections/test_b15_map_delete
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

let m = new Map();
        m.set("x", 10);
        m.set("y", 20);
        m.delete("x");
        __check(__line(m.has("x"), m.size), "false 1");
