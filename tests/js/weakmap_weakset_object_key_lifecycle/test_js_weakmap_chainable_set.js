// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakmap_chainable_set
// origin: languages/js/tests/js/test_js_weakmap_weakset_object_key_lifecycle.rs

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
const k1 = {}, k2 = {};
wm.set(k1, 1).set(k2, 2);
__check(__line(wm.get(k1) + "|" + wm.get(k2)), "1|2");
