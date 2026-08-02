// vybe-test: js/class_private_advanced/private_static_method_callable_on_class
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    static #double(x) { return x * 2; }
    static compute(x) { return MathHelper.#double(x) + 1; }
}
__check(__line(MathHelper.compute(5)), "11");
__check(__line(MathHelper.compute(10)), "21");
