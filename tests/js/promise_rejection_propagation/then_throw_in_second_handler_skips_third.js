// vybe-test: js/promise_rejection_propagation/then_throw_in_second_handler_skips_third
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

const o=[];Promise.resolve(1).then(v=>{o.push("a");return v;}).then(()=>{throw 1;}).then(()=>o.push("b")).catch(()=>o.push("c")).then(()=>console.log(o.join(",")));
