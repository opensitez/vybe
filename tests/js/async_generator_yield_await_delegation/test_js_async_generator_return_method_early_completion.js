// vybe-test: js/async_generator_yield_await_delegation/test_js_async_generator_return_method_early_completion
// origin: languages/js/tests/js/test_js_async_generator_yield_await_delegation.rs

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

async function* counter() {
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        __p(__line("Async Generator Cleanup"));
    }
}
(async () => {
    const gen = counter();
    __p(__line((await gen.next()).value));
    const r2 = await gen.return("EarlyReturn");
    __p(__line(`${r2.value}|done=${r2.done}`));
})();
__checkLater("1\nAsync Generator Cleanup\nEarlyReturn|done=true");
