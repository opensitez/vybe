// vybe-test: js/weakref_weakmap_advanced/weakset_delete_removes
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

const ws = new WeakSet();
const obj = {};
ws.add(obj);
ws.delete(obj);
__check(__line(ws.has(obj)), "false");
