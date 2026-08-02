// vybe-test: js/object_descriptors/from_entries_round_trips_with_entries
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

const original = { a: 1, b: 2, c: 3 };
const modified = Object.fromEntries(
    Object.entries(original).map(([k, v]) => [k, v * 2])
);
__check(__line(modified.a), "2");
__check(__line(modified.b), "4");
__check(__line(modified.c), "6");
