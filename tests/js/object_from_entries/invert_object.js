// vybe-test: js/object_from_entries/invert_object
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

const obj = { a: "1", b: "2", c: "3" };
const inverted = Object.fromEntries(
    Object.entries(obj).map(([k, v]) => [v, k])
);
__check(__line(inverted["1"]), "a");
__check(__line(inverted["2"]), "b");
