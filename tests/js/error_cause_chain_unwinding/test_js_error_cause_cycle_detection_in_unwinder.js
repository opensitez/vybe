// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_cycle_detection_in_unwinder
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

function getCauseChainSafe(err) {
    const visited = new Set();
    const chain = [];
    let current = err;
    while (current && !visited.has(current)) {
        visited.add(current);
        chain.push(current.message);
        current = current.cause;
    }
    return chain;
}
const e1 = new Error("E1");
const e2 = new Error("E2", { cause: e1 });
e1.cause = e2; // Cyclic cause chain!

console.log(getCauseChainSafe(e2).join(" -> "));
