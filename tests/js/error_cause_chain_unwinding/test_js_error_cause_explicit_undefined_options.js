// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_explicit_undefined_options
// origin: languages/js/tests/js/test_js_error_cause_chain_unwinding.rs

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

const err = new Error("Msg", { cause: undefined });
__check(__line(Object.hasOwn(err, "cause")), "true");
