// vybe-test: js/async_await_error_recovery/async_for_loop_break_on_caught_error
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

async function main(){const o=[];for(let i=0;i<5;i++){try{await Promise.reject(i);}catch(e){o.push(e);if(e===2)break;}}console.log(o.join(","));}main();
