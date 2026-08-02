// vybe-test: js/ecma_objects/object_spread
// origin: languages/js/tests/js/test_ecma_objects.rs

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

const a = { x: 1 };
const b = { y: 2 };
const merged = { ...a, ...b, z: 3 };
__check(__line(merged.x), "1");
__check(__line(merged.y), "2");
__check(__line(merged.z), "3");
