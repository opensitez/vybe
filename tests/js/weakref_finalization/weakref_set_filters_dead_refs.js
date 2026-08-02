// vybe-test: js/weakref_finalization/weakref_set_filters_dead_refs
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

// Simulate live reference tracking — objects still in scope are alive
let a = { id: "a" };
let b = { id: "b" };

const refs = [new WeakRef(a), new WeakRef(b)];
const live = refs.map(r => r.deref()).filter(Boolean);
__check(__line(live.length), "2");
__check(__line(live.map(o => o.id).join(",")), "a,b");
