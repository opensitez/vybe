// vybe-test: js/module_meta_loading/dynamic_import_exposes_namespace_keys
// origin: languages/js/tests/js/test_module_meta_loading.rs

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

const ns = await import("wasi:logging/logging");
console.log(Object.keys(ns).includes("log"));
