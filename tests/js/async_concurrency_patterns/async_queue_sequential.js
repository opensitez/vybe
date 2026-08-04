// vybe-test: js/async_concurrency_patterns/async_queue_sequential
// origin: languages/js/tests/js/test_async_concurrency_patterns.rs

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

class AsyncQueue {
    #queue = [];
    #running = false;
    enqueue(task) {
        return new Promise((resolve, reject) => {
            this.#queue.push({ task, resolve, reject });
            this.#run();
        });
    }
    async #run() {
        if (this.#running) return;
        this.#running = true;
        while (this.#queue.length) {
            const { task, resolve, reject } = this.#queue.shift();
            try { resolve(await task()); } catch(e) { reject(e); }
        }
        this.#running = false;
    }
}
async function main() {
    const q = new AsyncQueue();
    const order = [];
    const results = await Promise.all([
        q.enqueue(async () => { order.push(1); return "a"; }),
        q.enqueue(async () => { order.push(2); return "b"; }),
        q.enqueue(async () => { order.push(3); return "c"; }),
    ]);
    __p(__line(results.join(",")));
    __p(__line(order.join(",")));
}
main();
__checkLater("a,b,c\n1,2,3");
