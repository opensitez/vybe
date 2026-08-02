// vybe-test: js/ecma_error_handling/catch_binding_shadows_outer_variable
// origin: languages/js/tests/js/test_ecma_error_handling.rs

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

let e = "outer";
try {
    throw "inner";
} catch (e) {
    __check(__line(e), "inner");
}
__check(__line(e), "outer");
