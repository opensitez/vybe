// vybe-test: js/coercion_modern/test_boolean_coercion_primitive_object_wrappers_always_truthy
// origin: languages/js/tests/js/test_coercion_modern.rs

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

__check(__line(Boolean(new Boolean(false)) + "|" + Boolean(new Number(0)) + "|" + Boolean(new String(""))), "true|true|true");
