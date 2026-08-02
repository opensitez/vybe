// vybe-test: js/promise_rejection_propagation/catch_logs_and_returns_for_chain_continue
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

const o=[];Promise.reject("x").catch(e=>{o.push("log:"+e);return "go";}).then(v=>o.push(v)).then(()=>console.log(o.join(",")));
