// vybe-test: js/promise_finally_errors/finally_runs_after_catch_without_rethrow
// origin: languages/js/tests/js/test_promise_finally_errors.rs

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

const o=[];Promise.reject("x").catch(e=>o.push("c:"+e)).finally(()=>o.push("f")).then(()=>console.log(o.join(",")));
