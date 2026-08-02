// vybe-test: js/async_await_error_recovery/async_iife_immediate_invoke_with_await
// origin: languages/js/tests/js/test_async_await_error_recovery.rs

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

async function main(){const r=await(async()=>{try{await Promise.reject("q");}catch(e){return e+"!";}})();console.log(r);}main();
