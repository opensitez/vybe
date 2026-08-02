// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakmap_set_get_has_delete_flow
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
const key = { id: 1 };
wm.set(key, "PrivateData");

__check(__line(wm.get(key) + "|" + wm.has(key)), "PrivateData|true");
wm.delete(key);
__check(__line(wm.has(key)), "false");
