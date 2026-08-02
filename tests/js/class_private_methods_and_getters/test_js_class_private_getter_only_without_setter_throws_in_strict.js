// vybe-test: js/class_private_methods_and_getters/test_js_class_private_getter_only_without_setter_throws_in_strict
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

class ConstantHolder {
    get #constVal() { return "Immutable"; }
    trySet() {
        "use strict";
        try {
            this.#constVal = "NewVal";
        } catch (e) {
            __check(__line("Private Getter Set TypeError"), "Private Getter Set TypeError");
        }
    }
}
new ConstantHolder().trySet();
