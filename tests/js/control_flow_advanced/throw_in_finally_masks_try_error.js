// vybe-test: js/control_flow_advanced/throw_in_finally_masks_try_error
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

let caught;
try {
    try { throw new Error("original"); }
    finally { throw new Error("finally"); }
} catch (e) {
    caught = e.message;
}
__check(__line(caught), "finally");
