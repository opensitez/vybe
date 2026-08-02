// vybe-test: js/promise_rejection_propagation/then_throw_after_promise_all_member_reject
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

Promise.all([Promise.reject("m")]).catch(e=>e).then(v=>{throw "wrap:"+v;}).catch(e=>console.log(e));
