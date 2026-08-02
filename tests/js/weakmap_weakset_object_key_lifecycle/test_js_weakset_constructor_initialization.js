// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakset_constructor_initialization
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

const o1 = {}, o2 = {};
const ws = new WeakSet([o1, o2]);
__check(__line(ws.has(o1) + "|" + ws.has(o2)), "true|true");
