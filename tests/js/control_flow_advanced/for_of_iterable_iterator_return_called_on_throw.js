// vybe-test: js/control_flow_advanced/for_of_iterable_iterator_return_called_on_throw
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

let nextCount = 0;
let returnCount = 0;
const iterable = {
    [Symbol.iterator]() {
        return {
            next() {
                nextCount++;
                return nextCount === 1
                    ? { value: nextCount, done: false }
                    : { done: true };
            },
            return() {
                returnCount++;
                return { done: true };
            }
        };
    }
};

try {
    for (const value of iterable) {
        if (value === 1) {
            throw new Error("loop failure");
        }
    }
} catch (e) {
    __p(__line(e.message));
}
__p(__line(`${nextCount}:${returnCount}`));
__checkLater("loop failure\n1:1");
