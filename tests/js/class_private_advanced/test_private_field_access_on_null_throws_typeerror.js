// vybe-test: js/class_private_advanced/test_private_field_access_on_null_throws_typeerror
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    readNull(obj) {
        return obj.#code;
    }
}
const s = new Secret();
try {
    s.readNull(null);
} catch (e) {
    __check(__line(e.name), "TypeError");
}
