// vybe-test: js/weakref_weakmap_advanced/weakset_constructor_accepts_iterable
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

const o1 = {}, o2 = {}, o3 = {};
const ws = new WeakSet([o1, o2, o3]);
__check(__line(ws.has(o1)), "true");
__check(__line(ws.has(o2)), "true");
