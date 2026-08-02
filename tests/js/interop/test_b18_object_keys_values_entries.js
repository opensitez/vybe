// vybe-test: js/interop/test_b18_object_keys_values_entries
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

let obj = { x: 10 };
        let keys = Object.keys(obj);
        let vals = Object.values(obj);
        let entries = Object.entries(obj);
        __check(__line(keys.length, vals[0], entries[0][0], entries[0][1]), "1 10 x 10");
