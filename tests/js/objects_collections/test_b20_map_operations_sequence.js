// vybe-test: js/objects_collections/test_b20_map_operations_sequence
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
        m.set("a", 1);
        m.set("b", 2);
        m.set("c", 3);
        m.delete("b");
        m.set("d", 4);
        __check(__line(m.size, m.has("a"), m.has("b"), m.get("d")), "3 true false 4");
