// vybe-test: js/design_patterns/strategy_pattern
// origin: languages/js/tests/js/test_design_patterns.rs

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

class Sorter {
    constructor(strategy) { this.strategy = strategy; }
    sort(arr) { return this.strategy([...arr]); }
}
const ascending = arr => arr.sort((a, b) => a - b);
const descending = arr => arr.sort((a, b) => b - a);
const nums = [3, 1, 4, 1, 5, 9, 2, 6];
const asc = new Sorter(ascending);
const desc = new Sorter(descending);
__p(__line(asc.sort(nums).join(",")));
__p(__line(desc.sort(nums).join(",")));
__checkLater("1,1,2,3,4,5,6,9\n9,6,5,4,3,2,1,1");
