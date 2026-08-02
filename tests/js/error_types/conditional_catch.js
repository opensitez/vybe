// vybe-test: js/error_types/conditional_catch
// origin: languages/js/tests/js/test_error_types.rs

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

function process(val) {
    if (val < 0) throw new RangeError("negative");
    if (val === 0) throw new TypeError("zero not allowed");
    return val * 2;
}
try {
    process(-1);
} catch (e) {
    if (e instanceof RangeError) {
        console.log("range: " + e.message);
    } else if (e instanceof TypeError) {
        console.log("type: " + e.message);
    }
}
