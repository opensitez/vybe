// vybe-test: js/promise_finally_errors/finally_return_after_thenable_assimilation
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

Promise.resolve({then(res){res(5);}}).then(v=>v).finally(()=>"f").then(v=>console.log(v));
