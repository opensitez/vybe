// vybe-test: js/try_catch_finally_edge/finally_executes_even_without_error
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

const log = [];
try {
    log.push("try");
} catch {
    log.push("catch");
} finally {
    log.push("finally");
}
__check(__line(log.join(",")), "try,finally");
