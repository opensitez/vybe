// vybe-test: js/esm_host_imports/wasi_filesystem_preopens_actual_surface_namespace_import
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

import * as preopens from "wasi:filesystem/preopens";
const directories = preopens["get-directories"]();
__check(__line(Array.isArray(directories)), "true");
__check(__line(directories.length > 0), "true");
__check(__line(directories[0][1] === "."), "true");
