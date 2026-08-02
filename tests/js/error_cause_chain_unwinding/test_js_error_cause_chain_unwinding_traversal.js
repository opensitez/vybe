// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_chain_unwinding_traversal
// origin: languages/js/tests/js/test_js_error_cause_chain_unwinding.rs

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

function getCauseChain(err) {
    const chain = [];
    let current = err;
    while (current) {
        chain.push(current.message);
        current = current.cause;
    }
    return chain;
}
const dbErr = new Error("Database Connection Timeout");
const serviceErr = new Error("UserService Failed", { cause: dbErr });
const apiErr = new Error("HTTP 500 Internal Server Error", { cause: serviceErr });

console.log(getCauseChain(apiErr).join(" <- "));
