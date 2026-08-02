// vybe-test: js/interop/test_d37_static_methods
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
            static count = 0;
            static increment() { Counter.count++; return Counter.count; }
        }
        __check(__line(Counter.increment(), Counter.increment(), Counter.increment()), "1 2 3");
