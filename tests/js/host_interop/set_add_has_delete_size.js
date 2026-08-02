// vybe-test: js/host_interop/set_add_has_delete_size
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

let s = new Set();
        s.add(1);
        s.add(2);
        s.add(2);
        __check(__line(s.size), "2");
        __check(__line(s.has(1)), "true");
        s.delete(1);
        __check(__line(s.size), "1");
        __check(__line(s.has(1)), "false");
