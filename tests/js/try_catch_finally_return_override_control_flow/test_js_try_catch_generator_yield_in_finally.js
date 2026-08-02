// vybe-test: js/try_catch_finally_return_override_control_flow/test_js_try_catch_generator_yield_in_finally
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

function* gen() {
    try {
        yield 1;
    } finally {
        yield 2;
    }
}
const g = gen();
__check(__line(`${g.next().value}:${g.next().value}:${g.next().done}`), "1:2:true");
