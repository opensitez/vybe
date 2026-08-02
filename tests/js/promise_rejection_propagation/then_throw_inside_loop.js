// vybe-test: js/promise_rejection_propagation/then_throw_inside_loop
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

Promise.resolve(3).then(n=>{for(let i=0;i<n;i++){if(i===2)throw "loop";}}).catch(e=>console.log(e));
