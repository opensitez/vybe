// vybe-test: js/object_methods_deep/object_from_entries_basic
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

const entries = [["a", 1], ["b", 2], ["c", 3]];
const obj = Object.fromEntries(entries);
__check(__line(obj.a + "," + obj.b + "," + obj.c), "1,2,3");
