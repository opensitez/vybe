// vybe-test: js/objects_collections/test_f50_class_map_field
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

class Registry {
            constructor() { this.data = new Map(); }
            register(key, val) { this.data.set(key, val); }
            lookup(key) { return this.data.get(key); }
        }
        let r = new Registry();
        r.register("host", "localhost");
        __check(__line(r.lookup("host")), "localhost");
