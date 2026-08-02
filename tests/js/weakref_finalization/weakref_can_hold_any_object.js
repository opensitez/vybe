// vybe-test: js/weakref_finalization/weakref_can_hold_any_object
// origin: languages/js/tests/js/test_weakref_finalization.rs

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

const fn1 = () => "hi";
const arr = [1, 2, 3];
const map = new Map();

const refs = [new WeakRef(fn1), new WeakRef(arr), new WeakRef(map)];
__check(__line(typeof refs[0].deref()), "function");
__check(__line(Array.isArray(refs[1].deref())), "true");
__check(__line(refs[2].deref() instanceof Map), "true");
