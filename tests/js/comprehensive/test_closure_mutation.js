// vybe-test: js/comprehensive/test_closure_mutation
// origin: languages/js/tests/js/js_comprehensive_test.rs

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
            return { inc: () => { n++; return n; }, get: () => n };
        }
        let c = counter();
        c.inc();
        c.inc();
        c.inc();
        __check(__line(c.get()), "3");
