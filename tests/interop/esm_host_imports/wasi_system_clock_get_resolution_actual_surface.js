// vybe-test: interop/esm_host_imports/wasi_system_clock_get_resolution_actual_surface
// origin: languages/js/tests/js/test_esm_host_imports.rs

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

import * as systemClock from "wasi:clocks/system-clock";
// 0.3.1 declares `get-resolution: func() -> duration`, and `duration = u64`
// NANOSECONDS (`clocks/wit/types.wit`) — a bare number. The 0.2 interface this
// replaces answered a `{ seconds, nanoseconds }` record instead, so this is a
// change of SHAPE and not just of name. The third assertion is what pins that
// down: a record would satisfy neither, but without it a record that happened
// to coerce would slip through the first two.
// Kebab-case is not a JS identifier, so it is read the way
// `preopens["get-directories"]` is.
const resolution = systemClock["get-resolution"]();
__p(__line(typeof resolution === "number"));
__p(__line(resolution > 0));
__p(__line(resolution.seconds === undefined));
__checkLater("true\ntrue\ntrue");
