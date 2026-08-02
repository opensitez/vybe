// vybe-test: js/objects_collections/test_b17_map_string_keys_number_values
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
        m.set("score", 100);
        m.set("lives", 3);
        let total = m.get("score") + m.get("lives");
        __check(__line(total), "103");
