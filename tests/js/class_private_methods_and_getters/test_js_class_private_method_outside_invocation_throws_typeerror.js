// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_outside_invocation_throws_typeerror
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

class Service {
    #internalAction() { return "Internal"; }
}
const s = new Service();
try {
    eval("s.#internalAction()");
} catch (e) {
    __check(__line("Outside Private Method Call Error"), "Outside Private Method Call Error");
}
