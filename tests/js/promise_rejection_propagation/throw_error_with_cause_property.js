// vybe-test: js/promise_rejection_propagation/throw_error_with_cause_property
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

const c=new Error("cause");Promise.resolve(0).then(()=>{const e=new Error("main");e.cause=c;throw e;}).catch(e=>console.log(e.cause.message));
