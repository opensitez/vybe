// vybe-test: js/class_private_advanced/private_field_counter_pattern
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    #n = 0;
    inc() { this.#n++; return this; }
    dec() { this.#n--; return this; }
    reset() { this.#n = 0; return this; }
    value() { return this.#n; }
}
const c = new Counter();
c.inc().inc().inc().dec();
__check(__line(c.value()), "2");
c.reset();
__check(__line(c.value()), "0");
