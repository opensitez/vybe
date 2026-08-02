// vybe-test: js/class_static_deep/static_field_shared_across_instances
// origin: languages/js/tests/js/test_class_static_deep.rs

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
    static count = 0;
    constructor() { Counter.count++; }
}
new Counter();
new Counter();
new Counter();
__check(__line(Counter.count), "3");
