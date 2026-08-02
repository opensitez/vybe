// vybe-test: js/object_descriptors/from_entries_from_map
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

const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
const obj = Object.fromEntries(m);
__check(__line(obj.a), "1");
__check(__line(obj.b), "2");
__check(__line(obj.c), "3");
