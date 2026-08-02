// vybe-test: js/weakref_finalization/weakref_deref_with_nullish_coalescing
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

const obj = { name: "test" };
const ref1 = new WeakRef(obj);
const name = ref1.deref()?.name ?? "gone";
__check(__line(name), "test");
