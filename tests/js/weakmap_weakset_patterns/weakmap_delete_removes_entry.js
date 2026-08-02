// vybe-test: js/weakmap_weakset_patterns/weakmap_delete_removes_entry
// origin: languages/js/tests/js/test_weakmap_weakset_patterns.rs

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

const wm = new WeakMap();
const key = {};
wm.set(key, 42);
__check(__line(wm.has(key)), "true");
wm.delete(key);
__check(__line(wm.has(key)), "false");
