// vybe-test: js/async_await_error_recovery/async_loop_accumulator_survives_catch
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

async function main(){let sum=0;for(let i=1;i<=3;i++){try{sum+=await Promise.resolve(i);}catch{sum=-1;}if(i===2)try{await Promise.reject(0);}catch{}}console.log(sum);}main();
