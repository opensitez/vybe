// vybe-test: js/objects_collections/test_f54_class_iterate_array_field
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

class Summer {
            constructor() { this.values = []; }
            add(v) { this.values.push(v); }
            total() {
                let s = 0;
                let i = 0;
                while (i < this.values.length) {
                    s = s + this.values[i];
                    i = i + 1;
                }
                return s;
            }
        }
        let sm = new Summer();
        sm.add(10);
        sm.add(20);
        sm.add(30);
        console.log(sm.total());
