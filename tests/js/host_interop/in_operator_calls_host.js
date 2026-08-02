// vybe-test: js/host_interop/in_operator_calls_host
// origin: languages/js/tests/js/js_host_interop_test.rs

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

let obj = {a: 1, b: 2};
        __check(__line("a" in obj), "true");
        __check(__line("c" in obj), "false");
