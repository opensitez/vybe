// vybe-test: js/map_set/map_overwrite_no_size_change
// origin: languages/js/tests/js/js_map_set_test.rs

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
        m.set("a", 1);
        m.set("a", 2);
        __check(__line(m.size, m.get("a")), "1 2");
