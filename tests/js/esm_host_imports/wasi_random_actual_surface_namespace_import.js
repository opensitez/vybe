// vybe-test: js/esm_host_imports/wasi_random_actual_surface_namespace_import
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

import * as random from "wasi:random/random";
const bytes = random["get-random-bytes"](4);
const value = random["get-random-u64"]();
__check(__line(Array.isArray(bytes)), "true");
__check(__line(bytes.length === 4), "true");
__check(__line(typeof value === "number"), "true");
