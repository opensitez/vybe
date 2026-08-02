// vybe-test: js/spread_rest_advanced/object_spread_override_order
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

const base = { x: 1, y: 2, z: 3 };
const override = { x: 10 };
const merged = { ...base, ...override };
__check(__line(merged.x, merged.y), "10 2");
