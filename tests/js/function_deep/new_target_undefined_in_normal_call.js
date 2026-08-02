// vybe-test: js/function_deep/new_target_undefined_in_normal_call
// origin: languages/js/tests/js/test_function_deep.rs

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

function f() { return new.target; }
__check(__line(f() === undefined), "true");
__check(__line(new f() === undefined), "false");
