// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_recursion
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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
    #factorial(n) {
        if (n <= 1) return 1;
        return n * this.#factorial(n - 1);
    }
    fact(n) { return this.#factorial(n); }
}
__check(__line(new MathUtils().fact(5)), "120");
