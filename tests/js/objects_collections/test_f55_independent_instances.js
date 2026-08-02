// vybe-test: js/objects_collections/test_f55_independent_instances
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

class Stack {
            constructor() { this.items = []; }
            push(v) { this.items.push(v); }
            size() { return this.items.length; }
        }
        let a = new Stack();
        let b = new Stack();
        a.push(1);
        a.push(2);
        b.push(99);
        __check(__line(a.size(), b.size()), "2 1");
