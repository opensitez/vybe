// vybe-test: js/async_await_error_recovery/async_await_reject_after_successful_awaits_in_loop
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

async function main(){const o=[];for(let i=0;i<3;i++){try{o.push(await Promise.resolve(i));if(i===2)await Promise.reject("done");}catch(e){o.push("e:"+e);}}console.log(o.join(","));}main();
