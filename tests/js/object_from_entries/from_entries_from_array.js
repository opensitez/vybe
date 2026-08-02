// vybe-test: js/object_from_entries/from_entries_from_array
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

const entries = [["a", 1], ["b", 2], ["c", 3]];
const obj = Object.fromEntries(entries);
__check(__line(obj.a), "1");
__check(__line(obj.b), "2");
__check(__line(obj.c), "3");
