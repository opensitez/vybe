// vybe-test: js/map_set_advanced_patterns/map_with_complex_keys
// origin: languages/js/tests/js/test_map_set_advanced_patterns.rs

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

const keyA = { id: 1 };
const keyB = { id: 2 };
const m = new Map();
m.set(keyA, "first");
m.set(keyB, "second");
m.set({id:1}, "third"); // different object reference
__check(__line(m.size), "3");
__check(__line(m.get(keyA)), "first");
