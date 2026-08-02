// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_non_iterable_throws_typeerror
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
    structuredClone(123, { transfer: 12345 });
} catch (e) {
    __check(__line("Transfer Option Non-Iterable TypeError"), "Transfer Option Non-Iterable TypeError");
}
