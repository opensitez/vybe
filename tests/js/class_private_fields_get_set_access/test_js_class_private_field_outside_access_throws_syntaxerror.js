// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_outside_access_throws_syntaxerror
// origin: languages/js/tests/js/test_js_class_private_fields_get_set_access.rs

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

class Secret {
    #code = 1234;
}
const s = new Secret();
try {
    eval("s.#code");
} catch (e) {
    __check(__line("Outside Private Access Error"), "Outside Private Access Error");
}
