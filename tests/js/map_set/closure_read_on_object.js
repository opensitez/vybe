// vybe-test: js/map_set/closure_read_on_object
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

function make() { let x = 99; return { getValue: () => { return x; } }; }
        let o = make();
        __check(__line(o.getValue()), "99");
