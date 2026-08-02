// vybe-test: js/promise_rejection_propagation/then_rejection_skips_fulfilled_handler
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

const o=[];Promise.reject("x").then(()=>o.push("then")).catch(()=>o.push("catch")).then(()=>console.log(o.join(",")));
