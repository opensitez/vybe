// vybe-test: js/interop/test_g69_zero_falsy_string_zero_truthy
// origin: languages/js/tests/js/js_interop_test.rs

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

__check(__line(0 ? "truthy" : "falsy"), "falsy");
        __check(__line("0" ? "truthy" : "falsy"), "truthy");
