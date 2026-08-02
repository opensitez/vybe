// vybe-test: js/try_catch_finally_return_override_control_flow/test_js_try_finally_continue_loop
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
for (let i = 0; i < 2; i++) {
    try {
        log.push(`Try${i}`);
        continue;
    } finally {
        log.push(`Finally${i}`);
    }
}
console.log(log.join(","));
