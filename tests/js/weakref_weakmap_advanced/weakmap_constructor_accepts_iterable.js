// vybe-test: js/weakref_weakmap_advanced/weakmap_constructor_accepts_iterable
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

const k1 = {}, k2 = {};
const wm = new WeakMap([[k1, "a"], [k2, "b"]]);
__check(__line(wm.get(k1)), "a");
__check(__line(wm.get(k2)), "b");
