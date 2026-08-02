// vybe-test: js/missing_features/map_set_get_has
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

let m = new Map();
        m.set("key", "value");
        __check(__line(m.get("key")), "value");
        __check(__line(m.has("key")), "true");
        __check(__line(m.has("missing")), "false");
