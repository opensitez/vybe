// vybe-test: js/inheritance/test_11_static_method_on_base
// origin: languages/js/tests/js/js_inheritance_test.rs

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

class MathHelper {
            static add(a, b) { return a + b; }
        }
        __check(__line(MathHelper.add(2, 3)), "5");
