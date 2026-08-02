// vybe-test: js/weakref_weakmap_advanced/weakref_deref_returns_object
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

let obj = { value: 42 };
const ref = new WeakRef(obj);
const deref = ref.deref();
__check(__line(deref !== undefined), "true");
__check(__line(deref.value), "42");
