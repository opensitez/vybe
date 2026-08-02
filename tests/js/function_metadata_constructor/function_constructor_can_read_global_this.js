// vybe-test: js/function_metadata_constructor/function_constructor_can_read_global_this
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

globalThis.dynamicValue = 7;
const fn = new Function("return globalThis.dynamicValue;");
__check(__line(fn()), "7");
delete globalThis.dynamicValue;
