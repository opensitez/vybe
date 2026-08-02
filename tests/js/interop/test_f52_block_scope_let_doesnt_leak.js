// vybe-test: js/interop/test_f52_block_scope_let_doesnt_leak
// origin: languages/js/tests/js/js_interop_test.rs

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

let x = "outer";
        if (true) {
            let x = "inner";
            __check(__line(x), "inner");
        }
        __check(__line(x), "outer");
