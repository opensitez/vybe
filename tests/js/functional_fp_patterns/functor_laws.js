// vybe-test: js/functional_fp_patterns/functor_laws
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

// Output is COLLECTED, not paired. The emitter rewrites every `console.log(a)`
// into `__p(__line(a))` and compares the whole buffer once.
//
// Collection is what makes ASYNC assertable at all — 967 of the 1,860 cases the
// per-print emitter refused were `await` / `then` / `Promise`, where the i-th
// log in the SOURCE is not the i-th line of OUTPUT. The buffer records the
// order things actually ran, so no ordering analysis is needed.
let __buf = "";

function __p(s) {
    __buf += s + "\n";
}

function __pr(s) {
    __buf += s;
}

// The check runs from a `setTimeout(…, 0)` — a MACROtask, so it fires only
// after the microtask queue has fully drained. Measured under Vybe: a program
// logging sync, then a `.then`, then past an `await`, then the timeout,
// collects them in exactly that order, while a statement at the end of the
// script sees an empty buffer.
function __checkLater(want) {
    setTimeout(function () {
        __check(__buf, want);
    }, 0);
}

function __check(got, want) {
    // The final log contributes a trailing newline the expected line vector
    // never carried, so both forms are accepted.
    if (got !== want && got !== want + "\n") {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

class Box {
    constructor(v) { this._v = v; }
    map(fn) { return new Box(fn(this._v)); }
    value() { return this._v; }
}
const identity = x => x;
const double = x => x * 2;
const addOne = x => x + 1;
// Identity law: map(id) === id
__p(__line(new Box(5).map(identity).value() === new Box(5).value()));
// Composition law: map(f).map(g) === map(g(f(x)))
const a = new Box(5).map(double).map(addOne).value();
const b = new Box(5).map(x => addOne(double(x))).value();
__p(__line(a === b));
__checkLater("true\ntrue");
