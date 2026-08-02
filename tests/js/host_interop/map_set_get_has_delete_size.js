// vybe-test: js/host_interop/map_set_get_has_delete_size
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

let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        __check(__line(m.size), "2");
        __check(__line(m.get("a")), "1");
        __check(__line(m.has("b")), "true");
        m.delete("a");
        __check(__line(m.size), "1");
        __check(__line(m.has("a")), "false");
