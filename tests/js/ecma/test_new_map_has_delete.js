// vybe-test: js/ecma/test_new_map_has_delete
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

let m = new Map();
        m.set("x", 1);
        __check(__line(m.has("x")), "true");
        m.delete("x");
        __check(__line(m.has("x")), "false");
