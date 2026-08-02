// vybe-test: js/missing_features/set_delete
// origin: languages/js/tests/js/js_missing_features_test.rs

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
        s.add("a");
        s.add("b");
        s.delete("a");
        __check(__line(s.size), "1");
        __check(__line(s.has("a")), "false");
