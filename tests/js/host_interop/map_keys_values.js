// vybe-test: js/host_interop/map_keys_values
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
        let k = m.keys();
        let v = m.values();
        __check(__line(k.length), "undefined");
        __check(__line(v.length), "undefined");
