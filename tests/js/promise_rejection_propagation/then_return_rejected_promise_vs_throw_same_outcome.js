// vybe-test: js/promise_rejection_propagation/then_return_rejected_promise_vs_throw_same_outcome
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

const o=[];Promise.resolve(0).then(()=>Promise.reject("a")).catch(e=>o.push("1:"+e));Promise.resolve(0).then(()=>{throw "a";}).catch(e=>o.push("2:"+e)).then(()=>console.log(o.join("|")));
