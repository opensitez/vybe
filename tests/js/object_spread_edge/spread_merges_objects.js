// vybe-test: js/object_spread_edge/spread_merges_objects
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

const a = { x: 1, y: 2 };
const b = { y: 3, z: 4 };
const c = { ...a, ...b };
__check(__line(c.x), "1");
__check(__line(c.y), "3"); // b wins
__check(__line(c.z), "4");
