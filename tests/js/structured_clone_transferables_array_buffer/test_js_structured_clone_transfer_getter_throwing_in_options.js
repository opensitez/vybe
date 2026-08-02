// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_getter_throwing_in_options
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

const opts = {
    get transfer() { throw new Error("TransferGetterError"); }
};
try {
    structuredClone(123, opts);
} catch (e) {
    __check(__line(e.message), "TransferGetterError");
}
