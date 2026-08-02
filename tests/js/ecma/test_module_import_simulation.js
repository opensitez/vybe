// vybe-test: js/ecma/test_module_import_simulation
// origin: languages/js/tests/js/js_ecma_test.rs

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

// --- imported from utils.js ---
        export function capitalize(str) {
            if (str.length === 0) return str;
            return str.charAt(0).toUpperCase() + str.slice(1);
        }
        export let VERSION = "1.0.0";
        // --- main module ---
        __check(__line(capitalize("hello")), "Hello");
        __check(__line(VERSION), "1.0.0");
