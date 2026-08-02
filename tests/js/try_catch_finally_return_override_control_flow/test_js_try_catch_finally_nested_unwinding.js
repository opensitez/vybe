// vybe-test: js/try_catch_finally_return_override_control_flow/test_js_try_catch_finally_nested_unwinding
// origin: languages/js/tests/js/test_js_try_catch_finally_return_override_control_flow.rs

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
function fn() {
    try {
        log.push("Try1");
        try {
            log.push("Try2");
            throw new Error("Err2");
        } catch (e) {
            log.push("Catch2");
            return "Ret2";
        } finally {
            log.push("Finally2");
        }
    } finally {
        log.push("Finally1");
    }
}
__check(__line(fn() + "|" + log.join(",")), "Ret2|Try1,Try2,Catch2,Finally2,Finally1");
