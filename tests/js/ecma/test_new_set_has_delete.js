// vybe-test: js/ecma/test_new_set_has_delete
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

let s = new Set();
        s.add(42);
        __check(__line(s.has(42)), "true");
        s.delete(42);
        __check(__line(s.has(42)), "false");
