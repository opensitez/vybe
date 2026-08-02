// vybe-test: js/ecma_classes/class_static_method
// origin: languages/js/tests/js/test_ecma_classes.rs

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

class MathUtils {
    static square(x) { return x * x; }
    static cube(x) { return x * x * x; }
}
__check(__line(MathUtils.square(4)), "16");
__check(__line(MathUtils.cube(3)), "27");
