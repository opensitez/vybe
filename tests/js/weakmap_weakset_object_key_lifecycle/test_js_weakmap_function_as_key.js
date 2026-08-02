// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakmap_function_as_key
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
function fnKey() {}
wm.set(fnKey, "FunctionMetaData");
__check(__line(wm.get(fnKey)), "FunctionMetaData");
