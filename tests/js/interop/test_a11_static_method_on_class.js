// vybe-test: js/interop/test_a11_static_method_on_class
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

class MathUtil {
            static square(x) { return x * x; }
        }
        __check(__line(MathUtil.square(7)), "49");
