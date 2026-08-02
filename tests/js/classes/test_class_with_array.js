// vybe-test: js/classes/test_class_with_array
// origin: languages/js/tests/js/js_classes_test.rs

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
            constructor() {
                this.items = [];
            }
            push(item) {
                this.items.push(item);
            }
            size() {
                return this.items.length;
            }
        }
        let s = new Stack();
        s.push(1);
        s.push(2);
        s.push(3);
        __check(__line(s.items.length, s.size()), "3 3");
