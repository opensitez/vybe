// vybe-test: js/error_handling_advanced/error_cause_chain
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

function level3() { throw new Error("level3 fail"); }
function level2() {
    try { level3(); }
    catch(e) { throw new Error("level2 fail", { cause: e }); }
}
function level1() {
    try { level2(); }
    catch(e) { throw new Error("level1 fail", { cause: e }); }
}
try { level1(); } catch(e) {
    __check(__line(e.message), "level1 fail");
    __check(__line(e.cause.message), "level2 fail");
    __check(__line(e.cause.cause.message), "level3 fail");
}
