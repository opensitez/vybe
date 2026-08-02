// vybe-test: js/promise_rejection_propagation/catch_on_sync_throw_equivalent_to_reject
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

const o=[];Promise.resolve().then(()=>{throw "s";}).catch(e=>o.push("s:"+e));Promise.reject("a").catch(e=>o.push("a:"+e)).then(()=>console.log(o.sort().join("|")));
