// vybe-test: js/nullish_optional_deep/nullish_coalescing_zero_returns_left
// origin: languages/js/tests/js/test_nullish_optional_deep.rs

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

__check(__line(0 ?? "ignored"), "0");
__check(__line("" ?? "ignored"), "");
__check(__line(false ?? "ignored"), "false");
