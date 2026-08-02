// vybe-test: js/class_private_deep/private_field_per_instance
// origin: languages/js/tests/js/test_class_private_deep.rs

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
    inc() { this.#count++; }
    get() { return this.#count; }
}
const a = new Counter();
const b = new Counter();
a.inc(); a.inc();
b.inc();
__check(__line(a.get()), "2");
__check(__line(b.get()), "1");
