// vybe-test: js/class_inheritance_advanced/subclass_calls_super_method_with_this
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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
    increment() { this.count++; return this; }
}
class BoundedCounter extends Counter {
    constructor(max) {
        super();
        this.max = max;
    }
    increment() {
        if (this.count < this.max) super.increment();
        return this;
    }
}
const bc = new BoundedCounter(3);
bc.increment().increment().increment().increment().increment();
__check(__line(bc.count), "3");
