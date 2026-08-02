// vybe-test: js/missing_features/this_in_class_method
// origin: languages/js/tests/js/js_missing_features_test.rs

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
            constructor() { this.count = 0; }
            inc() { this.count++; }
            get() { return this.count; }
        }
        let c = new Counter();
        c.inc();
        c.inc();
        __check(__line(c.get()), "2");
