// vybe-test: js/promise_rejection_propagation/then_multiple_handlers_only_first_runs
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

const o=[];const p=Promise.resolve(1);p.then(v=>o.push("h1:"+v));p.then(v=>o.push("h2:"+v));p.then(()=>console.log(o.join(",")));
