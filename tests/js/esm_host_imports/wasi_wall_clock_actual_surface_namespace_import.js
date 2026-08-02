// vybe-test: js/esm_host_imports/wasi_wall_clock_actual_surface_namespace_import
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

import * as wallClock from "wasi:clocks/wall-clock";
const now = wallClock.now();
__check(__line(typeof now === "object"), "true");
__check(__line(now.seconds > 0), "true");
__check(__line(now.nanoseconds >= 0), "true");
