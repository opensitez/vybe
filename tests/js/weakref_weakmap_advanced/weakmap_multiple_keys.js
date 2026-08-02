// vybe-test: js/weakref_weakmap_advanced/weakmap_multiple_keys
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
const k1 = {}, k2 = {}, k3 = {};
wm.set(k1, 1);
wm.set(k2, 2);
wm.set(k3, 3);
__check(__line(wm.get(k1) + wm.get(k2) + wm.get(k3)), "6");
