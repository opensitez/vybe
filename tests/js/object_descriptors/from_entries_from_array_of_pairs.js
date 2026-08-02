// vybe-test: js/object_descriptors/from_entries_from_array_of_pairs
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

const entries = [["x", 10], ["y", 20]];
const obj = Object.fromEntries(entries);
__check(__line(obj.x + obj.y), "30");
