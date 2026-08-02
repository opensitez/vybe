// vybe-test: js/class_patterns/class_instance_fields_are_per_instance
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
    value = 0;
    inc() { this.value += 1; }
}
let a = new Counter();
let b = new Counter();
a.inc();
a.inc();
b.inc();
__check(__line(a.value), "2");
__check(__line(b.value), "1");
