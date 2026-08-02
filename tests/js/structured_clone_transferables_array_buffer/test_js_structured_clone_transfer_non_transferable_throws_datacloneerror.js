// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_non_transferable_throws_datacloneerror
// origin: languages/js/tests/js/test_js_structured_clone_transferables_array_buffer.rs

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

try {
    structuredClone({ a: 1 }, { transfer: [{ a: 1 }] });
} catch (e) {
    __check(__line("DataCloneError Non-Transferable"), "DataCloneError Non-Transferable");
}
