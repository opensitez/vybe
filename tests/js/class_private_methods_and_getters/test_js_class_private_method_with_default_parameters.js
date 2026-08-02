// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_with_default_parameters
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

class Helper {
    #format(val, prefix = "[INFO]") {
        return `${prefix} ${val}`;
    }
    info(msg) { return this.#format(msg); }
}
__check(__line(new Helper().info("System Ready")), "[INFO] System Ready");
