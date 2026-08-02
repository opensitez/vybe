// vybe-test: js/objects_collections/test_f56_class_shared_reference
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

class Modifier {
            constructor(obj) { this.obj = obj; }
            setVal(v) { this.obj.val = v; }
        }
        let shared = { val: 0 };
        let m1 = new Modifier(shared);
        let m2 = new Modifier(shared);
        m1.setVal(10);
        __check(__line(shared.val), "10");
        m2.setVal(20);
        __check(__line(shared.val), "20");
