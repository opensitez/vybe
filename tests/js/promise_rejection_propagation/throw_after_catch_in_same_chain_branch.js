// vybe-test: js/promise_rejection_propagation/throw_after_catch_in_same_chain_branch
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

Promise.reject("a").catch(()=>"b").then(v=>{if(v==="b")throw "c";}).catch(e=>console.log(e));
