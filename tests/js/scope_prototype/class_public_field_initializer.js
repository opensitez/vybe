// vybe-test: js/scope_prototype/class_public_field_initializer
// origin: languages/js/tests/js/test_scope_prototype.rs

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
    increment() { this.count++; }
}
let c = new Counter();
c.increment();
c.increment();
__check(__line(c.count), "2");
