// vybe-test: js/advanced/test_closure_mutation
// origin: languages/js/tests/js/js_advanced_test.rs

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

function counter() {
            let n = 0;
            return {
                inc() { n = n + 1; return n; },
                get() { return n; }
            };
        }
        // Object methods don't have 'this' bound, but closures work
        // We can't call c.inc() as a method yet, but we can test closure capture
        let n = 0;
        function inc() { n = n + 1; return n; }
        __check(__line(inc(), inc(), inc()), "1 2 3");
