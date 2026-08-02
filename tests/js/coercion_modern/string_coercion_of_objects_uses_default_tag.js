// vybe-test: js/coercion_modern/string_coercion_of_objects_uses_default_tag
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

__check(__line(String({})), "[object Object]");
__check(__line(String({ a: 1 })), "[object Object]");
