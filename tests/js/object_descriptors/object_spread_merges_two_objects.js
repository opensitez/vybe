// vybe-test: js/object_descriptors/object_spread_merges_two_objects
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

const a = { x: 1, y: 2 };
const b = { y: 99, z: 3 };
const merged = { ...a, ...b };
__check(__line(merged.x), "1");
__check(__line(merged.y), "99");
__check(__line(merged.z), "3");
