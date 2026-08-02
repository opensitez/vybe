// vybe-test: js/async_await_error_recovery/async_switch_await_rejection_in_case
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

async function main(){const o=[];for(const k of [1,2]){switch(k){case 1:try{await Promise.reject("s");}catch(e){o.push(e);}break;case 2:o.push("ok");}}console.log(o.join(","));}main();
