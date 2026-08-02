// vybe-test: js/class_patterns/test_extracted_class_method_bound_with_bind
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
    constructor() { this.value = 42; }
    get() { return this.value; }
}
const c = new Counter();
const boundGet = c.get.bind(c);
__check(__line(boundGet()), "42");
