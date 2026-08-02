// vybe-test: js/class_static_deep/static_method_called_on_class
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

class MathHelper {
    static add(a, b) { return a + b; }
    static multiply(a, b) { return a * b; }
}
__check(__line(MathHelper.add(3, 4)), "7");
__check(__line(MathHelper.multiply(3, 4)), "12");
