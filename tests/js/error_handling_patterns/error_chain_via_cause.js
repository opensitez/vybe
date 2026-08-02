// vybe-test: js/error_handling_patterns/error_chain_via_cause
// origin: languages/js/tests/js/test_error_handling_patterns.rs

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

function fetchData() {
    throw new TypeError("network error");
}
function loadUser(id) {
    try {
        return fetchData(id);
    } catch (e) {
        throw new Error("Failed to load user " + id, { cause: e });
    }
}
try {
    loadUser(42);
} catch (e) {
    __check(__line(e.message), "Failed to load user 42");
    __check(__line(e.cause instanceof TypeError), "true");
    __check(__line(e.cause.message), "network error");
}
