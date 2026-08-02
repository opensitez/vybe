// vybe-test: js/class_patterns/class_method_extracted_loses_this_binding
// origin: languages/js/tests/js/test_class_patterns.rs

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
    constructor() { this.value = 3; }
    get() { return this && this.value; }
}
let c = new Counter();
let fn = c.get;
__check(__line(c.get()), "3");
__check(__line(fn()), "undefined");
