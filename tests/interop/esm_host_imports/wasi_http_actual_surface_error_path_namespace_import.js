// vybe-test: interop/esm_host_imports/wasi_http_actual_surface_error_path_namespace_import
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

import * as httpTypes from "wasi:http/types";
import * as client from "wasi:http/client";
// 0.3.1 collapsed 0.2's `outgoing-request`/`incoming-request` pair into one
// `request` resource, and replaced the `outgoing-handler.handle` ->
// `future-incoming-response.get` two-step with a single `client.send` that
// answers the response directly. There is no future to consume here.
const headers = httpTypes["[constructor]fields"]();
const request = httpTypes["[static]request.new"](headers);
httpTypes["[method]request.set-scheme"](request, "http");
httpTypes["[method]request.set-authority"](request, "127.0.0.1:1");
httpTypes["[method]request.set-path-with-query"](request, "/");
const result = client.send(request);
__p(__line(result.__wasi_error === "connection-refused" || result.__wasi_error === "internal-error"));
__checkLater("true");
