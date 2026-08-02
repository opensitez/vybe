// vybe-test: js/objects_collections/test_a04_object_pass_by_reference
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

function modify(o) { o.val = 42; }
        let obj = { val: 0 };
        modify(obj);
        __check(__line(obj.val), "42");
