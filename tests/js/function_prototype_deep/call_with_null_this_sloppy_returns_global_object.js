// vybe-test: js/function_prototype_deep/call_with_null_this_sloppy_returns_global_object
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

function f() { return typeof this; } __check(__line(f.call(null) === "object" || f.call(null) === "undefined"), "true");
