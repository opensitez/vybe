// vybe-test: js/async_error_handling/async_rethrow_unknown_errors
// origin: languages/js/tests/js/test_async_error_handling.rs

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

class NetworkError extends Error {}

async function fetchSafe() {
    try {
        throw new NetworkError("timeout");
    } catch (e) {
        if (e instanceof NetworkError) return null;
        throw e; // rethrow unknown
    }
}

fetchSafe().then(v => console.log(v));
