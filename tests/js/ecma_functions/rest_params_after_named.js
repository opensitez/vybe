// vybe-test: js/ecma_functions/rest_params_after_named
// origin: languages/js/tests/js/test_ecma_functions.rs

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

function log(prefix, ...messages) {
    for (const m of messages) {
        console.log(prefix + ": " + m);
    }
}
log("INFO", "start", "end");
