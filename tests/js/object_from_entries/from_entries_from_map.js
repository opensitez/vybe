// vybe-test: js/object_from_entries/from_entries_from_map
// origin: languages/js/tests/js/test_object_from_entries.rs

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

const map = new Map([["x", 10], ["y", 20]]);
const obj = Object.fromEntries(map);
__check(__line(obj.x), "10");
__check(__line(obj.y), "20");
