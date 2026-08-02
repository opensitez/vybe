// vybe-test: js/esm_host_imports/node_os_platform_returns_node_faithful_string
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

import { platform } from "node:os";
let p = platform();
__check(__line(p === "darwin" || p === "linux" || p === "win32"), "true");
