// vybe-test: js/async_try_catch_finally_rejection_propagation/test_js_async_try_catch_with_custom_error_types
// origin: languages/js/tests/js/test_js_async_try_catch_finally_rejection_propagation.rs

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

class ValidationError extends Error {}
class NetworkError extends Error {}

async function process(type) {
    try {
        if (type === "val") throw new ValidationError("Invalid Input");
        if (type === "net") throw new NetworkError("Connection Timeout");
    } catch (e) {
        if (e instanceof ValidationError) console.log("ValError: " + e.message);
        else if (e instanceof NetworkError) console.log("NetError: " + e.message);
    }
}
(async () => {
    await process("val");
    await process("net");
})();
