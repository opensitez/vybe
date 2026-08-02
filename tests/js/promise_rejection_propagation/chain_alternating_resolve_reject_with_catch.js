// vybe-test: js/promise_rejection_propagation/chain_alternating_resolve_reject_with_catch
// origin: languages/js/tests/js/test_promise_rejection_propagation.rs

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

Promise.resolve(1).then(()=>Promise.reject("r")).catch(e=>e+"!").then(v=>console.log(v));
