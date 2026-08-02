// vybe-test: js/error_handling_advanced/optional_catch_binding
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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

function safeParse(s) {
    try { return { ok: true, val: JSON.parse(s) }; }
    catch { return { ok: false }; }
}
__check(__line(safeParse('{"x":1}').ok), "true");
__check(__line(safeParse("bad").ok), "false");
