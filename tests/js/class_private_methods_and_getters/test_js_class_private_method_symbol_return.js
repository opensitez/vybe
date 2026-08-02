// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_symbol_return
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

class KeyGen {
    #genSymbol(desc) { return Symbol(desc); }
    createKey(desc) {
        const s = this.#genSymbol(desc);
        return typeof s === "symbol";
    }
}
__check(__line(new KeyGen().createKey("auth")), "true");
