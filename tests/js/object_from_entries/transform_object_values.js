// vybe-test: js/object_from_entries/transform_object_values
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

const prices = { apple: 1.5, banana: 0.75, cherry: 2.0 };
const doubled = Object.fromEntries(
    Object.entries(prices).map(([k, v]) => [k, v * 2])
);
__check(__line(doubled.apple), "3");
__check(__line(doubled.banana), "1.5");
