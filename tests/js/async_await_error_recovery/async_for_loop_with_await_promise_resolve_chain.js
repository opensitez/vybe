// vybe-test: js/async_await_error_recovery/async_for_loop_with_await_promise_resolve_chain
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

async function main(){const o=[];for(let i=0;i<2;i++){o.push(await Promise.resolve(i).then(x=>x+1));}console.log(o.join(","));}main();
