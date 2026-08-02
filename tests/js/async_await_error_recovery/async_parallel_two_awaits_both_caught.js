// vybe-test: js/async_await_error_recovery/async_parallel_two_awaits_both_caught
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

async function main(){const o=[];const a=(async()=>{try{await Promise.reject("1");}catch(e){o.push(e);}})();const b=(async()=>{try{await Promise.reject("2");}catch(e){o.push(e);}})();await Promise.all([a,b]);console.log(o.sort().join(","));}main();
