// vybe-test: js/reactive_patterns/event_emitter_pattern
// origin: languages/js/tests/js/test_reactive_patterns.rs

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

class EventEmitter {
    #listeners = new Map();
    on(event, fn) {
        if (!this.#listeners.has(event)) this.#listeners.set(event, []);
        this.#listeners.get(event).push(fn);
        return this;
    }
    off(event, fn) {
        const list = this.#listeners.get(event) || [];
        this.#listeners.set(event, list.filter(f => f !== fn));
        return this;
    }
    emit(event, ...args) {
        (this.#listeners.get(event) || []).forEach(fn => fn(...args));
        return this;
    }
    once(event, fn) {
        const wrapper = (...args) => { fn(...args); this.off(event, wrapper); };
        return this.on(event, wrapper);
    }
}
const ee = new EventEmitter();
const results = [];
ee.on("data", v => results.push("on:" + v));
ee.once("data", v => results.push("once:" + v));
ee.emit("data", 1);
ee.emit("data", 2);
__p(__line(results.join(",")));
__checkLater("on:1,once:1,on:2");
