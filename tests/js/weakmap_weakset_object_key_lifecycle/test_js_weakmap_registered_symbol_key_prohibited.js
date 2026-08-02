// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakmap_registered_symbol_key_prohibited
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
const globalSym = Symbol.for("globalKey");
try {
    wm.set(globalSym, "Val");
} catch (e) {
    console.log("WeakMap Global Registered Symbol TypeError");
}
