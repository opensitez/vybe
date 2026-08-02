// vybe-test: js/objects_collections/test_e41_function_modifies_object
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

function inc(obj) { obj.count = obj.count + 1; }
        let o = { count: 0 };
        inc(o);
        inc(o);
        inc(o);
        __check(__line(o.count), "3");
