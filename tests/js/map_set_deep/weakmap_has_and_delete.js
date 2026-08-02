// vybe-test: js/map_set_deep/weakmap_has_and_delete
// origin: languages/js/tests/js/test_map_set_deep.rs

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
wm.set(key, "value");
__check(__line(wm.has(key)), "true");
wm.delete(key);
__check(__line(wm.has(key)), "false");
