// vybe-test: js/esm_host_imports/wasi_insecure_seed_actual_surface_namespace_import
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

import * as seed from "wasi:random/insecure-seed";
const pair = seed["insecure-seed"]();
__check(__line(Array.isArray(pair)), "true");
__check(__line(pair.length === 2), "true");
__check(__line(typeof pair[0] === "number"), "true");
