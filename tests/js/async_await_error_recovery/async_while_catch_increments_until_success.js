// vybe-test: js/async_await_error_recovery/async_while_catch_increments_until_success
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

async function main(){let n=0;while(n<3){try{if(n<2)await Promise.reject(n);else break;}catch{n++;}}console.log(n);}main();
