// vybe-test: js/interop/test_f56_nested_closures_three_levels
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

function a() {
            let x = 1;
            function b() {
                let y = 2;
                function c() {
                    let z = 3;
                    return x + y + z;
                }
                return c();
            }
            return b();
        }
        __check(__line(a()), "6");
