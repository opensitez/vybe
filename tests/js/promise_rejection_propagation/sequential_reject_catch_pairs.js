// vybe-test: js/promise_rejection_propagation/sequential_reject_catch_pairs
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

const o=[];Promise.reject(1).catch(e=>o.push("a:"+e));Promise.reject(2).catch(e=>o.push("b:"+e));Promise.resolve().then(()=>console.log(o.join(",")));
