// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_field_outside_access_throws
// origin: languages/js/tests/js/test_js_class_static_private_fields_methods.rs

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

class Vault {
    static #pass = 9999;
}
try {
    eval("Vault.#pass");
} catch (e) {
    __check(__line("Outside Static Private Access Error"), "Outside Static Private Access Error");
}
