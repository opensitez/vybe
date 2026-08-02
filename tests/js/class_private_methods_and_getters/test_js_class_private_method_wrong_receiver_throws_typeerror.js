// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_wrong_receiver_throws_typeerror
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

class Component {
    #render() { return "DOM"; }
    callRender(other) {
        return other.#render();
    }
}
const c = new Component();
try {
    c.callRender({});
} catch (e) {
    __check(__line("Private Method Wrong Receiver TypeError"), "Private Method Wrong Receiver TypeError");
}
