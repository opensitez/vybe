// vybe-test: js/object_from_entries/entries_then_from_entries_roundtrip
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

const original = { a: 1, b: 2, c: 3 };
const clone = Object.fromEntries(Object.entries(original));
__check(__line(clone.a), "1");
__check(__line(clone.b), "2");
__check(__line(clone.c), "3");
