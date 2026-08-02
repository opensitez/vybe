// vybe-test: js/ecma/test_set_timeout_async
// origin: languages/js/tests/js/js_ecma_test.rs

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

let log = "";
        setTimeout(() => { log = log + "timer "; console.log(log + "done"); }, 1);
        log = log + "sync ";
        console.log(log);
