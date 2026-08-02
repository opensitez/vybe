// vybe-test: js/weakref_finalization/weakref_deref_returns_object_while_alive
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

let obj = { value: 42 };
const ref1 = new WeakRef(obj);
__check(__line(ref1.deref()?.value), "42");
