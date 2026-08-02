// vybe-test: js/class_patterns/static_block_computed
// origin: languages/js/tests/js/test_class_patterns.rs

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

class MathConstants {
    static PI;
    static TAU;
    static {
        MathConstants.PI = 3.14159;
        MathConstants.TAU = MathConstants.PI * 2;
    }
}
__check(__line(MathConstants.PI), "3.14159");
__check(__line(MathConstants.TAU), "6.28318");
