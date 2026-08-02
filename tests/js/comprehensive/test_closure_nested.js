// vybe-test: js/comprehensive/test_closure_nested
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

function a() {
            let x = 1;
            function b() {
                let y = 2;
                return () => x + y;
            }
            return b();
        }
        __check(__line(a()()), "3");
