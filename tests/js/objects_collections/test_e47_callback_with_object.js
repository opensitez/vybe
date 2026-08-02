// vybe-test: js/objects_collections/test_e47_callback_with_object
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

function process(obj, cb) { cb(obj); }
        let data = { val: 10 };
        process(data, function(o) { o.val = o.val * 2; });
        __check(__line(data.val), "20");
