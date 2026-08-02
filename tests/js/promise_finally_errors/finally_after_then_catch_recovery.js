// vybe-test: js/promise_finally_errors/finally_after_then_catch_recovery
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

Promise.resolve(0).then(()=>{throw "t";}).catch(()=>"ok").finally(()=>{throw "f";}).catch(e=>console.log(e));
