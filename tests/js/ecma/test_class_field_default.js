// vybe-test: js/ecma/test_class_field_default
// origin: languages/js/tests/js/js_ecma_test.rs

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
            count = 0;
            increment() { this.count = this.count + 1; }
            getCount() { return this.count; }
        }
        let c = new Counter();
        c.increment();
        c.increment();
        __check(__line(c.getCount()), "2");
