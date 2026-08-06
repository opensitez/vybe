// vybe-test: js/async_concurrency_patterns/async_semaphore
// origin: languages/js/tests/js/test_async_concurrency_patterns.rs

function __fmt(v) {
    // console.log renders a bigint with an `n` suffix; String() drops it.
    return typeof v === "bigint" ? String(v) + "n" : String(v);
}

function __line(...args) {
    // console.log joins its arguments with a single space. __fmt is the
    // per-argument coercion console.log applies.
    return args.map(__fmt).join(" ");
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

class Semaphore {
    #count;
    #queue = [];
    constructor(count) { this.#count = count; }
    async acquire() {
        if (this.#count > 0) { this.#count--; return; }
        await new Promise(resolve => this.#queue.push(resolve));
    }
    release() {
        if (this.#queue.length) { this.#queue.shift()(); }
        else { this.#count++; }
    }
}
async function main() {
    const sem = new Semaphore(2);
    const log = [];
    async function task(id) {
        await sem.acquire();
        log.push("in:" + id);
        await Promise.resolve();
        log.push("out:" + id);
        sem.release();
    }
    await Promise.all([task(1), task(2), task(3)]);
    __p(__line(log.includes("in:1")));
    __p(__line(log.includes("in:2")));
    __p(__line(log.includes("out:3")));
}
main();
__checkLater("true\ntrue\ntrue");
