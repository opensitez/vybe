// vybe-test: js/promise_rejection_propagation/then_fulfilled_handler_not_called_after_throw
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

const o=[];Promise.resolve(1).then(()=>{throw 1;}).then(()=>o.push("nope")).catch(()=>o.push("yes")).then(()=>console.log(o.join(",")));
