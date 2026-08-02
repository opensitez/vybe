// vybe-test: js/ecma_classes/class_private_field
// origin: languages/js/tests/js/test_ecma_classes.rs

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
    #count = 0;
    increment() { this.#count++; }
    get value() { return this.#count; }
}
const c = new Counter();
c.increment();
c.increment();
c.increment();
__check(__line(c.value), "3");
