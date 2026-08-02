// vybe-test: js/operators_deep/nullish_coalescing_on_boolean_and_zero
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line((false ?? "fallback")), "false");
__check(__line((0 ?? "fallback")), "0");
__check(__line((null ?? "fallback")), "fallback");
