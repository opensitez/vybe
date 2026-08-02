// vybe-test: js/function_metadata_constructor/function_constructor_does_not_capture_local_scope
// origin: languages/js/tests/js/test_function_metadata_constructor.rs

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

let hidden = 42;
const fn = new Function("return typeof hidden;");
__check(__line(fn()), "undefined");
