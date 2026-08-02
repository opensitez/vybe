// vybe-test: js/error_handling_advanced/error_propagation_through_callbacks
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

function safe(fn) {
    try { return { ok: true, value: fn() }; }
    catch(e) { return { ok: false, error: e.message }; }
}
const r1 = safe(() => JSON.parse('{"x":1}'));
const r2 = safe(() => JSON.parse("invalid"));
__check(__line(r1.ok), "true");
__check(__line(r1.value.x), "1");
__check(__line(r2.ok), "false");
__check(__line(typeof r2.error), "string");
