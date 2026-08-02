// vybe-test: js/weakref_weakmap_advanced/weakmap_has_returns_correct
// origin: languages/js/tests/js/test_weakref_weakmap_advanced.rs

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
const k1 = {};
const k2 = {};
wm.set(k1, 1);
__check(__line(wm.has(k1)), "true");
__check(__line(wm.has(k2)), "false");
