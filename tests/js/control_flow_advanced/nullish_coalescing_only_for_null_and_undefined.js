// vybe-test: js/control_flow_advanced/nullish_coalescing_only_for_null_and_undefined
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

__check(__line(0 ?? "fallback"), "0");
__check(__line("" ?? "fallback"), "");
__check(__line(false ?? "fallback"), "false");
__check(__line(null ?? "fallback"), "fallback");
__check(__line(undefined ?? "fallback"), "fallback");
