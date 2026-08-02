// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_rest_parameters
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

class Multiplier {
    #product(...nums) {
        return nums.reduce((a, b) => a * b, 1);
    }
    compute(...args) { return this.#product(...args); }
}
__check(__line(new Multiplier().compute(2, 3, 4)), "24");
