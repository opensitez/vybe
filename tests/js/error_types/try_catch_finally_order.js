// vybe-test: js/error_types/try_catch_finally_order
// origin: languages/js/tests/js/test_error_types.rs

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
    log.push("try");
    throw new Error("oops");
} catch (e) {
    log.push("catch");
} finally {
    log.push("finally");
}
__check(__line(log.join(",")), "try,catch,finally");
