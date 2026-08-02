// vybe-test: js/try_catch_finally_edge/finally_runs_after_throw
// origin: languages/js/tests/js/test_try_catch_finally_edge.rs

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

let ran = false;
function f() {
    try {
        throw new Error("boom");
    } finally {
        ran = true;
    }
}
try { f(); } catch {}
__check(__line(ran), "true");
