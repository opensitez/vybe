// vybe-test: js/scope_prototype/weakref_deref_returns_target
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let target = { id: 7 };
let wr = new WeakRef(target);
__check(__line(wr.deref().id), "7");
__check(__line(wr.deref() === target), "true");
