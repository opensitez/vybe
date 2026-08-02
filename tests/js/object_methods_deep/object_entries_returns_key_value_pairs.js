// vybe-test: js/object_methods_deep/object_entries_returns_key_value_pairs
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const obj = { a: 1, b: 2 };
const entries = Object.entries(obj);
__check(__line(entries.map(([k,v]) => k+"="+v).sort().join(",")), "a=1,b=2");
