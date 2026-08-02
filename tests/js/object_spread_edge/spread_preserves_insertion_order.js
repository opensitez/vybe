// vybe-test: js/object_spread_edge/spread_preserves_insertion_order
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

const defaults = { x: 1, y: 2, z: 3 };
const overrides = { ...defaults, y: 99 };
__check(__line(overrides.x), "1");
__check(__line(overrides.y), "99");
__check(__line(overrides.z), "3");
