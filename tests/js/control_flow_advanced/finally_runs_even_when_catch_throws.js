// vybe-test: js/control_flow_advanced/finally_runs_even_when_catch_throws
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

let log = [];
try {
    try {
        throw new Error("inner");
    } catch (e) {
        log.push("caught:" + e.message);
        throw new Error("rethrown");
    } finally {
        log.push("finally");
    }
} catch (e) {
    log.push("outer:" + e.message);
}
__check(__line(log.join("|")), "caught:inner|finally|outer:rethrown");
