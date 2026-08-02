// vybe-test: js/promise_rejection_propagation/rejection_skips_intermediate_thens_until_catch
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

const o=[];Promise.reject(0).then(()=>o.push("t1")).then(()=>o.push("t2")).catch(()=>o.push("c")).then(()=>console.log(o.join(",")));
