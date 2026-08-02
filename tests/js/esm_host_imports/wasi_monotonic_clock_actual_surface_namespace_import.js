// vybe-test: js/esm_host_imports/wasi_monotonic_clock_actual_surface_namespace_import
// origin: languages/js/tests/js/test_esm_host_imports.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

import * as monotonicClock from "wasi:clocks/monotonic-clock";
const now = monotonicClock.now();
__check(__line(typeof now === "number"), "true");
__check(__line(now >= 0), "true");
__check(__line(monotonicClock.resolution() >= 0), "true");
