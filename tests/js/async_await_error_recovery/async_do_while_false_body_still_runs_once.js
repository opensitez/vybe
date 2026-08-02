// vybe-test: js/async_await_error_recovery/async_do_while_false_body_still_runs_once
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

async function main(){let ran=false;do{try{await Promise.reject("once");}catch{ran=true;}}while(false);console.log(ran);}main();
