// vybe-test: js/async_await_error_recovery/async_parallel_await_all_individual_try_catch
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

async function main(){const ps=[Promise.resolve(1),Promise.reject("b"),Promise.resolve(3)];const o=[];for(const p of ps){try{o.push(await p);}catch(e){o.push("e:"+e);}}console.log(o.join(","));}main();
