// vybe-test: js/weakref_weakmap_advanced/weakref_target_not_collected_while_reachable
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

const target = { id: 99 };
const ref = new WeakRef(target);
const derefed = ref.deref();
__check(__line(derefed.id), "99");
