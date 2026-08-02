// vybe-test: js/esm_host_imports/wasi_cli_environment_actual_surface_namespace_import
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

import * as environment from "wasi:cli/environment";
const envPairs = environment["get-environment"]();
const args = environment["get-arguments"]();
const cwd = environment["initial-cwd"]();
__check(__line(Array.isArray(envPairs)), "true");
__check(__line(Array.isArray(args)), "true");
__check(__line(cwd === null || typeof cwd === "string"), "true");
