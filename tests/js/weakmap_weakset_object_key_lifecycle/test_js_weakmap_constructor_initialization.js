// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakmap_constructor_initialization
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

const k1 = {}, k2 = {};
const wm = new WeakMap([[k1, "Val1"], [k2, "Val2"]]);
__check(__line(wm.get(k1) + "|" + wm.get(k2)), "Val1|Val2");
