// vybe-test: js/function_prototype_deep/call_on_function_prototype_call_invokes_target
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

function target(v) { return v * 2; } __check(__line(Function.prototype.call.call(target, null, 4)), "8");
