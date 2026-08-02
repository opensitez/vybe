// vybe-test: js/objects_collections/test_f49_class_array_field_push
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

class Bag {
            constructor() { this.items = []; }
            add(item) { this.items.push(item); }
            count() { return this.items.length; }
        }
        let b = new Bag();
        b.add("apple");
        b.add("banana");
        __check(__line(b.count()), "2");
