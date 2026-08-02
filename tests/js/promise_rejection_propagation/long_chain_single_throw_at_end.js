// vybe-test: js/promise_rejection_propagation/long_chain_single_throw_at_end
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

Promise.resolve(1).then(x=>x+1).then(x=>x*2).then(x=>x-1).then(()=>{throw "end";}).catch(e=>console.log(e));
