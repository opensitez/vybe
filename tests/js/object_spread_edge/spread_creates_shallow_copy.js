// vybe-test: js/object_spread_edge/spread_creates_shallow_copy
// origin: languages/js/tests/js/test_object_spread_edge.rs

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

const original = { a: { x: 1 } };
const copy = { ...original };
copy.a.x = 99;
__check(__line(original.a.x), "99"); // shared reference
