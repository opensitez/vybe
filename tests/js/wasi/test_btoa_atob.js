// vybe-test: js/wasi/test_btoa_atob
// origin: languages/js/tests/js/js_wasi_test.rs

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

let encoded = btoa("Hello World");
        __check(__line(encoded), "SGVsbG8gV29ybGQ=");
        let decoded = atob(encoded);
        __check(__line(decoded), "Hello World");
