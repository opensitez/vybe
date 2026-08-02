// vybe-test: js/esm_host_imports/wildcard_namespace_includes_value_and_function_exports
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

import * as processNs from "node:process";
__check(__line(Array.isArray(processNs.argv)), "true");
__check(__line(typeof processNs.cwd === "function"), "true");
__check(__line(typeof processNs.version === "string"), "true");
__check(__line(typeof processNs.env === "object"), "true");
