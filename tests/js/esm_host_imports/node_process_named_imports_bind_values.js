// vybe-test: js/esm_host_imports/node_process_named_imports_bind_values
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

import { argv, env, versions, platform, arch, version, pid, execPath } from "node:process";
__check(__line(Array.isArray(argv)), "true");
__check(__line(typeof env === "object"), "true");
__check(__line(typeof versions.node === "string"), "true");
__check(__line(typeof platform === "string"), "true");
__check(__line(typeof arch === "string"), "true");
__check(__line(typeof version === "string"), "true");
__check(__line(typeof pid === "number"), "true");
__check(__line(typeof execPath === "string"), "true");
