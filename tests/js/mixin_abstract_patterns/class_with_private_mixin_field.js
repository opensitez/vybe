// vybe-test: js/mixin_abstract_patterns/class_with_private_mixin_field
// origin: languages/js/tests/js/test_mixin_abstract_patterns.rs

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
    getCount() { return this.#count; }
}
const c = new Counter();
c.increment();
c.increment();
__check(__line(c.getCount()), "2");
