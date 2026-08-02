// vybe-test: js/module_import_patterns/dynamic_import_failure_caught_by_catch
// origin: languages/js/tests/js/test_module_import_patterns.rs

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

import("./does_not_exist_xyz.js")
    .then(() => console.log("loaded"))
    .catch(() => console.log("failed"));
