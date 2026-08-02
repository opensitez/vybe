// vybe-test: js/async_await_error_recovery/async_while_await_increment_in_catch
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

async function main(){let i=0,o=[];while(i<2){try{await Promise.reject(i);}catch(e){o.push(e);i++;}}console.log(o.join(","));}main();
