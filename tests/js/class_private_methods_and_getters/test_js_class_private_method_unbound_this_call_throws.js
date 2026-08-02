// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_unbound_this_call_throws
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

class Detached {
    #method() { return "DetachedResult"; }
    getDetached() {
        const fn = this.#method;
        return fn(); // Calling detached private method without 'this' receiver throws TypeError!
    }
}
try {
    new Detached().getDetached();
} catch (e) {
    __check(__line("Detached Private Call TypeError"), "Detached Private Call TypeError");
}
