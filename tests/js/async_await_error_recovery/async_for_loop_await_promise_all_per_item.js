// vybe-test: js/async_await_error_recovery/async_for_loop_await_promise_all_per_item
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

async function main(){const o=[];for(const x of[1,2]){try{const[v]=await Promise.all([Promise.resolve(x)]);o.push(v);}catch(e){o.push(e);}}console.log(o.join(","));}main();
