// vybe-test: js/classes/test_class_multiple_methods
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

class Counter {
            constructor(start) {
                this.count = start;
            }
            increment() {
                this.count = this.count + 1;
            }
            get() {
                return this.count;
            }
        }
        let c = new Counter(0);
        c.increment();
        c.increment();
        c.increment();
        __check(__line(c.get()), "3");
