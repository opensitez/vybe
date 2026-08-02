// vybe-test: js/promise_finally_errors/finally_throw_after_catch_reraise
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

Promise.reject("a").catch(e=>{throw "b:"+e;}).finally(()=>{throw "f";}).catch(e=>console.log(e));
