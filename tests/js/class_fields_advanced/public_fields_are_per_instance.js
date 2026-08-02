// vybe-test: js/class_fields_advanced/public_fields_are_per_instance
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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
const a = new Counter();
const b = new Counter();
a.increment(); a.increment();
b.increment();
__check(__line(a.count), "2");
__check(__line(b.count), "1");
