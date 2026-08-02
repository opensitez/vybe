// vybe-test: js/object_descriptors/object_entries_returns_key_value_pairs
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const obj = { x: 1, y: 2 };
const entries = Object.entries(obj).sort(([a], [b]) => a < b ? -1 : 1);
__check(__line(entries.map(([k, v]) => k + ":" + v).join(",")), "x:1,y:2");
