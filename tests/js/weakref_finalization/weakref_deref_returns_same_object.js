// vybe-test: js/weakref_finalization/weakref_deref_returns_same_object
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

const obj = { id: 1 };
const ref1 = new WeakRef(obj);
__check(__line(ref1.deref() === obj), "true");
