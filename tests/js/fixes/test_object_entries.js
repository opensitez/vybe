// vybe-test: js/fixes/test_object_entries
// origin: languages/js/tests/js/js_fixes_test.rs

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

let obj = { x: 1 };
        let entries = Object.entries(obj);
        __check(__line(entries.length), "1");
