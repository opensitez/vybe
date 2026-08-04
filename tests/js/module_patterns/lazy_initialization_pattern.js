// vybe-test: js/module_patterns/lazy_initialization_pattern
// origin: languages/js/tests/js/test_module_patterns.rs

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

class LazyInit {
    #instance = null;
    #factory;
    constructor(factory) { this.#factory = factory; }
    get() {
        if (!this.#instance) {
            this.#instance = this.#factory();
        }
        return this.#instance;
    }
}
let created = 0;
const lazy = new LazyInit(() => { created++; return { value: 42 }; });
__p(__line(created));     // not yet created
const v = lazy.get();
__p(__line(created));     // now created
const v2 = lazy.get();
__p(__line(created));     // not created again
__p(__line(v === v2));    // same instance
__checkLater("0\n1\n1\ntrue");
