// vybe-test: js/promise_finally_errors/finally_throw_in_loop_inside_callback
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

Promise.resolve(0).finally(()=>{for(let i=0;i<2;i++){if(i===1)throw "loop";}}).catch(e=>console.log(e));
