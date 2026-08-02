// vybe-test: js/interop/test_d42_multiple_instances_independent_state
// origin: languages/js/tests/js/js_interop_test.rs

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
            constructor(start) { this.count = start; }
            inc() { this.count++; return this.count; }
        }
        let a = new Counter(0);
        let b = new Counter(100);
        a.inc(); a.inc(); a.inc();
        b.inc();
        __check(__line(a.count, b.count), "3 101");
