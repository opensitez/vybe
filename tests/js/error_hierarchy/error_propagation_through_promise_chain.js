// vybe-test: js/error_hierarchy/error_propagation_through_promise_chain
// origin: languages/js/tests/js/test_error_hierarchy.rs

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

class ApiError extends Error {
    constructor(msg, code) { super(msg); this.code = code; }
}

async function fetchData() {
    throw new ApiError("not found", 404);
}

fetchData().catch(e => {
    __check(__line(e instanceof ApiError), "true");
    __check(__line(e.code), "404");
    __check(__line(e.message), "not found");
});
