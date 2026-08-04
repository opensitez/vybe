// vybe-test: js/oop_patterns_advanced/chain_of_responsibility
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

class Handler {
    constructor(next = null) { this.next = next; }
    handle(req) { return this.next ? this.next.handle(req) : "unhandled"; }
}
class AuthHandler extends Handler {
    handle(req) { return req.auth ? super.handle(req) : "unauthorized"; }
}
class RateLimitHandler extends Handler {
    handle(req) { return req.rate > 100 ? "rate limited" : super.handle(req); }
}
class ResourceHandler extends Handler {
    handle(req) { return "ok:" + req.resource; }
}
const chain = new AuthHandler(new RateLimitHandler(new ResourceHandler()));
__p(__line(chain.handle({ auth: true, rate: 50, resource: "data" })));
__p(__line(chain.handle({ auth: false, rate: 50, resource: "data" })));
__p(__line(chain.handle({ auth: true, rate: 200, resource: "data" })));
__checkLater("ok:data\nunauthorized\nrate limited");
