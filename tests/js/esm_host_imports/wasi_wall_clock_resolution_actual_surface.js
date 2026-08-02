// vybe-test: js/esm_host_imports/wasi_wall_clock_resolution_actual_surface
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
const resolution = wallClock.resolution();
__check(__line(typeof resolution === "object"), "true");
__check(__line(resolution.seconds === 0), "true");
__check(__line(resolution.nanoseconds >= 0), "true");
