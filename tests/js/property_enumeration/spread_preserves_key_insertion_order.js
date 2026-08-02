// vybe-test: js/property_enumeration/spread_preserves_key_insertion_order
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const base = { x: 1, y: 2 };
const merged = { ...base, z: 3, x: 99 }; // x overridden
const keys = Object.keys(merged).sort();
__check(__line(keys.join(",")), "x,y,z");
__check(__line(merged.x), "99");
