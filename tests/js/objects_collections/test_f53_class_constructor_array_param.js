// vybe-test: js/objects_collections/test_f53_class_constructor_array_param
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

class Holder {
            constructor(data) { this.data = data; }
            first() { return this.data[0]; }
            size() { return this.data.length; }
        }
        let h = new Holder([10, 20, 30]);
        __check(__line(h.first(), h.size()), "10 3");
