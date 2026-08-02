// vybe-test: js/ecma/test_class_method_kinds
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
            constructor() {
                this._count = 0;
            }
            increment() { this._count = this._count + 1; }
            getCount() { return this._count; }
        }
        let c = new Counter();
        c.increment();
        c.increment();
        c.increment();
        __check(__line(c.getCount()), "3");
